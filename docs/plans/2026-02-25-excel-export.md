# Excel Export Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Add `.xlsx` workbook export (Summary + Repayment Schedule + embedded charts) to FinCalAgent — triggered from CLI (saves file to disk) and via a new `POST /export` HTTP endpoint.

**Architecture:** A new `src/excel-export.js` module receives raw FinCalEngine JSON and builds an ExcelJS workbook, then uses `jszip` to inject chart XML into the `.xlsx` archive after writing. `agent.js` captures the last FinCalEngine JSON on the agent instance; CLI and server layer detect export intent and call the export module with that captured data.

**Tech Stack:** Node.js ESM, `exceljs` (workbook/sheet/cell), `jszip` (zip manipulation for chart XML injection), existing `@anthropic-ai/sdk` + `just-bash` stack unchanged.

---

## FinCalEngine JSON shape (reference for all tasks)

The FinCalEngine API returns `data` with this structure (from `calcRepayments`):

```js
{
  amort_schedule: [
    {
      "Installment No": 1,          // 0 = disbursement row, skip it
      "Date": "2026-03-20",
      "Opening Balance": 5000000,
      "Installment Amount": 74405.10,
      "Interest": 62500,
      "Principal": 11905.10,
      "Closing Balance": 4988094.90,
      "Fees": 0,
      "Subvention": 0
    },
    // ... one row per period (plus row 0 disbursement)
  ],
  totals: [{ "Installment Amount": 10714334.40, "Interest": 5714334.40, "Principal": 5000000, "Fees": 0, "Subvention": 0 }],
  summary: [
    { "Step-Type": "Disbursement", "Date From": "...", "Date To": "...", "Installment Amount": 0 },
    { "Step-Type": "Step",         "Date From": "...", "Date To": "...", "Installment Amount": 74405.10 }
  ],
  returns: { "IRR": "15.00 %", "NDIRR": "...", "XIRR": "16.08 %", "NDXIRR": "..." }
}
```

The agent stores this in `this.lastLoanData` (added in Task 3). The loan params (principal, rate, term) are inside `agent.history` messages but are NOT needed for export — the schedule contains everything.

---

## Task 1: Install dependencies

**Files:**
- Modify: `package.json` (auto-updated by npm)

**Step 1: Install packages**

```bash
cd /Users/mohanjoshi/Documents/fincalc/fincal-agent2
npm install exceljs jszip
```

Expected output: `added N packages` with no errors.

**Step 2: Verify**

```bash
node -e "import('exceljs').then(m => console.log('exceljs ok')); import('jszip').then(m => console.log('jszip ok'))"
```

Expected: two lines printed, no errors.

**Step 3: Commit**

```bash
git add package.json package-lock.json
git commit -m "chore: add exceljs and jszip for Excel export"
```

---

## Task 2: Create `src/excel-export.js` — Summary + Schedule sheets

**Files:**
- Create: `src/excel-export.js`

**Step 1: Write the integration test first (TDD)**

Add this test to the bottom of `src/demo.js`, before the `main().catch(...)` call:

```js
// ── Excel export smoke test ──────────────────────────────────────────────────
import { buildWorkbook, writeToFile } from "./excel-export.js";
import { existsSync, statSync, unlinkSync } from "fs";

async function testExcelExport() {
  console.log(`\n${"─".repeat(60)}`);
  console.log("TEST: Excel export — Summary + Schedule sheets");
  console.log("─".repeat(60));

  // Minimal fixture — same shape as FinCalEngine JSON
  const fixture = {
    amort_schedule: [
      { "Installment No": 0, "Date": "2026-02-25", "Opening Balance": 100000, "Installment Amount": 0,       "Interest": 0,    "Principal": 0,      "Closing Balance": 100000, "Fees": 0, "Subvention": 0 },
      { "Installment No": 1, "Date": "2026-03-25", "Opening Balance": 100000, "Installment Amount": 9000.50, "Interest": 1250, "Principal": 7750.50, "Closing Balance": 92249.50, "Fees": 0, "Subvention": 0 },
      { "Installment No": 2, "Date": "2026-04-25", "Opening Balance": 92249.50,"Installment Amount": 9000.50,"Interest": 1153.12,"Principal":7847.38, "Closing Balance": 84402.12, "Fees": 0, "Subvention": 0 },
    ],
    totals:  [{ "Installment Amount": 108006, "Interest": 8006, "Principal": 100000, "Fees": 0, "Subvention": 0 }],
    summary: [
      { "Step-Type": "Disbursement", "Date From": "2026-02-25", "Date To": "2026-02-25", "Installment Amount": 0 },
      { "Step-Type": "Step",         "Date From": "2026-03-25", "Date To": "2027-02-25", "Installment Amount": 9000.50 },
    ],
    returns: { "IRR": "15.00 %", "NDIRR": "14.50 %", "XIRR": "16.08 %", "NDXIRR": "15.55 %" },
    loan:    { principal: 100000, rate: "15", termMths: 12 },
  };

  const wb   = await buildWorkbook(fixture);
  const path = "/tmp/fincal-test-export.xlsx";

  await writeToFile(wb, path);

  if (!existsSync(path))       throw new Error("File not created");
  if (statSync(path).size < 1000) throw new Error("File too small — likely empty");

  // Check sheet names
  const sheetNames = wb.worksheets.map(s => s.name);
  if (!sheetNames.includes("Summary"))  throw new Error("Missing Summary sheet");
  if (!sheetNames.includes("Schedule")) throw new Error("Missing Schedule sheet");
  if (!sheetNames.includes("Charts"))   throw new Error("Missing Charts sheet");

  unlinkSync(path);
  console.log("✅ Excel export test passed");
}
```

Also add `await testExcelExport();` inside the `main()` function, at the bottom before `console.log("\n✅ All tests complete.")`.

**Step 2: Run the test to confirm it fails**

```bash
node src/demo.js 2>&1 | tail -20
```

Expected: `Cannot find module './excel-export.js'` or similar import error.

**Step 3: Create `src/excel-export.js`**

```js
// src/excel-export.js
// Builds an Excel workbook from FinCalEngine JSON.
// Three sheets: Summary, Schedule, Charts.
// Chart XML is injected into the .xlsx zip archive after ExcelJS writes it.

import ExcelJS from "exceljs";
import JSZip   from "jszip";
import { readFileSync, writeFileSync, renameSync, unlinkSync } from "fs";
import { tmpdir } from "os";
import { join }   from "path";

// ── Colours ──────────────────────────────────────────────────────────────────
const C = {
  headerBg:   "FF1F2937",   // dark slate
  headerFont: "FFFFFFFF",   // white
  accentBg:   "FFF0F9FF",   // light blue tint
  posNum:     "FF166534",   // dark green for numbers
  border:     { style: "thin", color: { argb: "FFD1D5DB" } },
};

function borderAll() {
  return { top: C.border, left: C.border, bottom: C.border, right: C.border };
}

function inr(n) {
  return Number(n).toLocaleString("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// ── Sheet 1: Summary ─────────────────────────────────────────────────────────
function buildSummarySheet(wb, data) {
  const ws = wb.addWorksheet("Summary");
  ws.columns = [
    { width: 28 },
    { width: 24 },
  ];

  // Title
  ws.mergeCells("A1:B1");
  const title = ws.getCell("A1");
  title.value = "Loan Summary";
  title.font  = { bold: true, size: 14, color: { argb: C.headerFont } };
  title.fill  = { type: "pattern", pattern: "solid", fgColor: { argb: C.headerBg } };
  title.alignment = { horizontal: "center", vertical: "middle" };
  ws.getRow(1).height = 28;

  const tot  = data.totals[0];
  const ret  = data.returns;
  const sched = data.amort_schedule.filter(r => r["Installment No"] > 0);
  const first = sched[0];

  // Pull loan params if present (stored by agent.js)
  const loan = data.loan ?? {};

  const rows = [
    ["Principal",       loan.principal ? `₹${inr(loan.principal)}` : "—"],
    ["Interest Rate",   loan.rate      ? `${loan.rate}% p.a.`       : "—"],
    ["Term",            loan.termMths  ? `${loan.termMths} months`   : "—"],
    ["─────────────────────────", ""],
    ["First EMI",       `₹${inr(first?.["Installment Amount"] ?? 0)}`],
    ["Total Repayment", `₹${inr(tot["Installment Amount"])}`],
    ["Total Interest",  `₹${inr(tot["Interest"])}`],
    ["Total Principal", `₹${inr(tot["Principal"])}`],
    ...(tot["Fees"] > 0       ? [["Total Fees",      `₹${inr(tot["Fees"])}`]]       : []),
    ...(tot["Subvention"] > 0 ? [["Total Subvention",`₹${inr(tot["Subvention"])}`]] : []),
    ["─────────────────────────", ""],
    ["IRR",   ret["IRR"]],
    ["NDIRR", ret["NDIRR"]],
    ["XIRR",  ret["XIRR"]],
    ["NDXIRR",ret["NDXIRR"]],
  ];

  rows.forEach((pair, i) => {
    const rowNum = i + 2;
    const row    = ws.getRow(rowNum);
    row.getCell(1).value = pair[0];
    row.getCell(2).value = pair[1];

    if (pair[0].startsWith("─")) {
      row.getCell(1).font = { color: { argb: "FFD1D5DB" } };
      return;
    }

    row.getCell(1).font = { bold: true };
    row.getCell(2).font = { color: { argb: C.posNum } };

    if (i % 2 === 0) {
      [1, 2].forEach(col => {
        row.getCell(col).fill = { type: "pattern", pattern: "solid", fgColor: { argb: C.accentBg } };
      });
    }

    [1, 2].forEach(col => { row.getCell(col).border = borderAll(); });
  });
}

// ── Sheet 2: Schedule ─────────────────────────────────────────────────────────
function buildScheduleSheet(wb, data) {
  const ws = wb.addWorksheet("Schedule");

  const headers = ["Period", "Date", "Opening Balance", "EMI", "Interest", "Principal", "Closing Balance"];
  ws.columns = [
    { width: 8  },
    { width: 14 },
    { width: 20 },
    { width: 18 },
    { width: 16 },
    { width: 16 },
    { width: 20 },
  ];

  // Header row
  const hRow = ws.getRow(1);
  headers.forEach((h, i) => {
    const cell = hRow.getCell(i + 1);
    cell.value = h;
    cell.font  = { bold: true, color: { argb: C.headerFont } };
    cell.fill  = { type: "pattern", pattern: "solid", fgColor: { argb: C.headerBg } };
    cell.alignment = { horizontal: "center" };
    cell.border = borderAll();
  });
  hRow.height = 20;

  // Data rows (skip Installment No === 0 disbursement row)
  const sched = data.amort_schedule.filter(r => r["Installment No"] > 0);

  sched.forEach((r, i) => {
    const rowNum = i + 2;
    const row    = ws.getRow(rowNum);
    const vals   = [
      r["Installment No"],
      r["Date"],
      r["Opening Balance"],
      r["Installment Amount"],
      r["Interest"],
      r["Principal"],
      r["Closing Balance"],
    ];

    vals.forEach((v, col) => {
      const cell = row.getCell(col + 1);
      cell.value  = v;
      cell.border = borderAll();
      if (col >= 2) {
        cell.numFmt    = '#,##0.00';
        cell.alignment = { horizontal: "right" };
        cell.font      = { color: { argb: C.posNum } };
      }
    });

    if (i % 2 === 0) {
      [3, 4, 5, 6, 7].forEach(col => {
        row.getCell(col).fill = { type: "pattern", pattern: "solid", fgColor: { argb: C.accentBg } };
      });
    }
  });

  ws.autoFilter = { from: "A1", to: "G1" };
  ws.views = [{ state: "frozen", ySplit: 1 }];
}

// ── Sheet 3: Charts (placeholder — chart XML injected after write) ─────────────
function buildChartsSheet(wb) {
  const ws = wb.addWorksheet("Charts");
  ws.getCell("A1").value = "Charts are embedded in this sheet.";
  ws.getCell("A1").font  = { italic: true, color: { argb: "FF6B7280" } };
}

// ── Public: buildWorkbook ─────────────────────────────────────────────────────
export async function buildWorkbook(loanData) {
  if (!loanData?.amort_schedule) throw new Error("Invalid loan data: missing amort_schedule");
  if (!loanData?.totals?.length)  throw new Error("Invalid loan data: missing totals");
  if (!loanData?.returns)         throw new Error("Invalid loan data: missing returns");

  const wb = new ExcelJS.Workbook();
  wb.creator  = "FinCalAgent";
  wb.created  = new Date();

  buildSummarySheet(wb, loanData);
  buildScheduleSheet(wb, loanData);
  buildChartsSheet(wb);

  return wb;
}

// ── Public: writeToFile ───────────────────────────────────────────────────────
// Writes to a temp path first, injects charts, then renames to final path.
export async function writeToFile(workbook, filepath) {
  const tmp = join(tmpdir(), `fincal-${Date.now()}.xlsx`);
  await workbook.xlsx.writeFile(tmp);

  try {
    const withCharts = await injectCharts(tmp, workbook);
    writeFileSync(tmp + ".final", withCharts);
    renameSync(tmp + ".final", filepath);
  } catch (e) {
    // Graceful degradation: save without charts
    console.warn("[excel-export] Chart injection failed:", e.message, "— saving without charts");
    renameSync(tmp, filepath);
    return;
  }

  try { unlinkSync(tmp); } catch { /* already renamed */ }
}

// ── Public: toBuffer ──────────────────────────────────────────────────────────
export async function toBuffer(workbook) {
  const buf = await workbook.xlsx.writeBuffer();

  try {
    return await injectChartsToBuffer(Buffer.from(buf), workbook);
  } catch (e) {
    console.warn("[excel-export] Chart injection failed:", e.message, "— returning without charts");
    return Buffer.from(buf);
  }
}

// ── Chart XML injection (implemented in Task 4) ───────────────────────────────
// Stubs — replaced in Task 4
async function injectCharts(filepath, workbook) {
  const raw = readFileSync(filepath);
  return raw; // no-op until Task 4
}

async function injectChartsToBuffer(buf, workbook) {
  return buf; // no-op until Task 4
}
```

**Step 4: Run the test**

```bash
node src/demo.js 2>&1 | tail -10
```

Expected: `✅ Excel export test passed` followed by `✅ All tests complete.`

**Step 5: Commit**

```bash
git add src/excel-export.js src/demo.js
git commit -m "feat: add excel-export module with Summary and Schedule sheets"
```

---

## Task 3: Capture raw loan JSON in `agent.js`

The CLI and HTTP server need access to the raw FinCalEngine JSON after each `agent.ask()` call. We add `this.lastLoanData` to the agent and populate it whenever a bash tool result contains FinCalEngine JSON.

**Files:**
- Modify: `src/agent.js:150-207`

**Step 1: Add `lastLoanData` to constructor (line ~152)**

Find this line:
```js
  constructor({ verbose = false } = {}) {
    this.verbose = verbose;
    this.history = [];
  }
```

Replace with:
```js
  constructor({ verbose = false } = {}) {
    this.verbose  = verbose;
    this.history  = [];
    this.lastLoanData = null;   // populated after each successful FinCalEngine call
  }
```

**Step 2: Capture JSON in the tool_use loop (line ~182)**

Find this block inside the `tool_use` handler:
```js
          const { stdout, stderr, exitCode } = await execBash(script, this.verbose);

          toolResults.push({
```

Replace with:
```js
          const { stdout, stderr, exitCode } = await execBash(script, this.verbose);

          // Capture raw FinCalEngine JSON for Excel export
          if (exitCode === 0 && stdout.trim().startsWith("{")) {
            try {
              const parsed = JSON.parse(stdout.trim());
              if (parsed.amort_schedule && parsed.totals && parsed.returns) {
                this.lastLoanData = parsed;
              }
            } catch { /* not JSON, ignore */ }
          }

          toolResults.push({
```

**Step 3: Reset `lastLoanData` in `reset()`**

Find:
```js
  reset() { this.history = []; }
```

Replace with:
```js
  reset() { this.history = []; this.lastLoanData = null; }
```

**Step 4: Manual verify**

No automated test needed here — the CLI integration test in Task 5 will exercise this.

**Step 5: Commit**

```bash
git add src/agent.js
git commit -m "feat: capture lastLoanData on agent for Excel export"
```

---

## Task 4: Inject chart XML into `.xlsx` archive

An `.xlsx` file is a zip. After ExcelJS writes it, we use `jszip` to add chart XML files and update the relationships.

**Files:**
- Modify: `src/excel-export.js` (replace the two stub functions at the bottom)

**Background — xlsx chart XML structure:**

A chart in an xlsx lives in:
```
xl/charts/chart1.xml          ← chart definition
xl/drawings/drawing1.xml      ← anchor: positions the chart on the sheet
xl/drawings/_rels/drawing1.xml.rels  ← links drawing to chart
xl/worksheets/_rels/sheet3.xml.rels  ← links Charts sheet to drawing
[Content_Types].xml           ← must list the new files
```

**Step 1: Add chart XML template functions to `src/excel-export.js`**

Add these functions before the `injectCharts` stub (replace the stubs entirely):

```js
// ── Chart XML builders ────────────────────────────────────────────────────────

function chartXmlStackedBar(scheduleRowCount) {
  const lastRow = scheduleRowCount + 1; // +1 for header row
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    <c:title>
      <c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Principal vs Interest per Period</a:t></a:r></a:p></c:rich></c:tx>
      <c:overlay val="0"/>
    </c:title>
    <c:autoTitleDeleted val="0"/>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="stacked"/>
        <c:varyColors val="0"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:tx><c:strRef><c:f>Schedule!$E$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Interest</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:cat><c:numRef><c:f>Schedule!$A$2:$A$${lastRow}</c:f></c:numRef></c:cat>
          <c:val><c:numRef><c:f>Schedule!$E$2:$E$${lastRow}</c:f></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="1"/><c:order val="1"/>
          <c:tx><c:strRef><c:f>Schedule!$F$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Principal</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:cat><c:numRef><c:f>Schedule!$A$2:$A$${lastRow}</c:f></c:numRef></c:cat>
          <c:val><c:numRef><c:f>Schedule!$F$2:$F$${lastRow}</c:f></c:numRef></c:val>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:barChart>
      <c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
      <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
    </c:plotArea>
    <c:plotVisOnly val="1"/>
  </c:chart>
</c:chartSpace>`;
}

function chartXmlLineBalance(scheduleRowCount) {
  const lastRow = scheduleRowCount + 1;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    <c:title>
      <c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Outstanding Balance</a:t></a:r></a:p></c:rich></c:tx>
      <c:overlay val="0"/>
    </c:title>
    <c:autoTitleDeleted val="0"/>
    <c:plotArea>
      <c:lineChart>
        <c:grouping val="standard"/>
        <c:varyColors val="0"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:tx><c:strRef><c:f>Schedule!$G$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Closing Balance</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:cat><c:numRef><c:f>Schedule!$A$2:$A$${lastRow}</c:f></c:numRef></c:cat>
          <c:val><c:numRef><c:f>Schedule!$G$2:$G$${lastRow}</c:f></c:numRef></c:val>
        </c:ser>
        <c:axId val="3"/><c:axId val="4"/>
      </c:lineChart>
      <c:catAx><c:axId val="3"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="4"/></c:catAx>
      <c:valAx><c:axId val="4"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="3"/></c:valAx>
    </c:plotArea>
    <c:plotVisOnly val="1"/>
  </c:chart>
</c:chartSpace>`;
}

// Anchors chart in the Charts sheet at a fixed cell range
function drawingXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
           xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor moveWithCells="1" sizeWithCells="1">
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>9</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>15</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="2" name="Chart 1"/>
        <xdr:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></xdr:cNvGraphicFramePr>
      </xdr:nvGraphicFramePr>
      <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
        <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId1"/>
      </a:graphicData></a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
  <xdr:twoCellAnchor moveWithCells="1" sizeWithCells="1">
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>17</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>9</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>32</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="3" name="Chart 2"/>
        <xdr:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></xdr:cNvGraphicFramePr>
      </xdr:nvGraphicFramePr>
      <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
        <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId2"/>
      </a:graphicData></a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>`;
}

function drawingRelsXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart2.xml"/>
</Relationships>`;
}

// ── Core injection logic ──────────────────────────────────────────────────────

async function doInject(inputBuf, workbook) {
  const scheduleRowCount = workbook
    .getWorksheet("Schedule")
    .rowCount - 1; // subtract header

  const zip = await JSZip.loadAsync(inputBuf);

  // Find which sheet index is "Charts" (sheets are 1-indexed in xl/worksheets/sheetN.xml)
  let chartsSheetIdx = 0;
  workbook.worksheets.forEach((ws, i) => { if (ws.name === "Charts") chartsSheetIdx = i + 1; });
  if (!chartsSheetIdx) throw new Error("Charts worksheet not found in workbook");

  const sheetRelsPath  = `xl/worksheets/_rels/sheet${chartsSheetIdx}.xml.rels`;
  const drawingPath    = `xl/drawings/drawing1.xml`;
  const drawingRels    = `xl/drawings/_rels/drawing1.xml.rels`;
  const chart1Path     = `xl/charts/chart1.xml`;
  const chart2Path     = `xl/charts/chart2.xml`;

  // Add chart XML files
  zip.file(chart1Path, chartXmlStackedBar(scheduleRowCount));
  zip.file(chart2Path, chartXmlLineBalance(scheduleRowCount));

  // Add drawing (positions both charts on the Charts sheet)
  zip.file(drawingPath, drawingXml());
  zip.file(drawingRels, drawingRelsXml());

  // Link the Charts sheet to the drawing via its .rels file
  const sheetRel = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>`;
  zip.file(sheetRelsPath, sheetRel);

  // Update [Content_Types].xml to register new parts
  const ctRaw = await zip.file("[Content_Types].xml").async("string");
  const chartCt  = `<Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
  <Override PartName="/xl/charts/chart2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`;
  const ctUpdated = ctRaw.replace("</Types>", `  ${chartCt}\n</Types>`);
  zip.file("[Content_Types].xml", ctUpdated);

  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}

async function injectCharts(filepath, workbook) {
  const raw = readFileSync(filepath);
  return doInject(raw, workbook);
}

async function injectChartsToBuffer(buf, workbook) {
  return doInject(buf, workbook);
}
```

**Step 2: Run the integration test again**

```bash
node src/demo.js 2>&1 | tail -10
```

Expected: still `✅ Excel export test passed` — the chart injection shouldn't break the file structure.

**Step 3: Manual smoke test**

Run a real demo first to get real data, then manually export:

```bash
node -e "
import('./src/fincal-api.js').then(async ({ calcRepayments, buildPayload }) => {
  const data = await calcRepayments(buildPayload({
    principal: 1000000, rate: '15', termMths: 60,
    startDate: '2026-02-25', emiDate: '2026-03-25',
    repayStructure: [{ 'Step/Balloon/CashFlow': 'Step', 'No of Periods': 60, GrwthRate: 0, Value: 0 }]
  }));
  data.loan = { principal: 1000000, rate: '15', termMths: 60 };
  const { buildWorkbook, writeToFile } = await import('./src/excel-export.js');
  const wb = await buildWorkbook(data);
  await writeToFile(wb, '/tmp/fincal-smoke.xlsx');
  console.log('Saved /tmp/fincal-smoke.xlsx');
});
"
```

Open `/tmp/fincal-smoke.xlsx` in Excel, Numbers, or LibreOffice. Verify:
- Three sheets: Summary, Schedule, Charts
- Charts sheet shows two charts (stacked bar + line)
- Summary has key metrics
- Schedule has all repayment rows

**Step 4: Commit**

```bash
git add src/excel-export.js
git commit -m "feat: inject chart XML into xlsx — stacked bar and line charts"
```

---

## Task 5: CLI integration in `src/index.js`

**Files:**
- Modify: `src/index.js`

**Step 1: Add export intent detection and export call**

Find the import at the top of `src/index.js`:
```js
import { createInterface } from "readline";
import { FinCalAgent } from "./agent.js";
```

Replace with:
```js
import { createInterface } from "readline";
import { join }            from "path";
import { FinCalAgent }     from "./agent.js";
import { buildWorkbook, writeToFile } from "./excel-export.js";
```

Find the prompt handler body (the `try` block inside `rl.question`):
```js
    try {
      const answer = await agent.ask(q);
      console.log("Agent:", answer, "\n");
    } catch (e) {
      console.error("Error:", e.message, "\n");
    }
```

Replace with:
```js
    const wantsExcel = /export|xlsx|excel/i.test(q);

    try {
      const answer = await agent.ask(q);
      console.log("Agent:", answer, "\n");

      if (wantsExcel && agent.lastLoanData) {
        const filename = `loan-export-${new Date().toISOString().slice(0,10)}.xlsx`;
        const filepath = join(process.cwd(), filename);
        console.log("Generating Excel workbook...");
        const wb = await buildWorkbook(agent.lastLoanData);
        await writeToFile(wb, filepath);
        console.log(`Saved: ${filename}\n`);
      } else if (wantsExcel) {
        console.log("[No loan data available — calculate a loan first, then export]\n");
      }
    } catch (e) {
      console.error("Error:", e.message, "\n");
    }
```

Also update the help text to mention the export command. Find:
```js
  console.log("  Try:  Compare flat vs step-up structure");
```

Add after it:
```js
  console.log("  Try:  Export to Excel");
```

**Step 2: Manual test**

```bash
ANTHROPIC_API_KEY=<your-key> node src/index.js
```

Type: `EMI for 10 Lakh at 15% for 5 years, export to Excel`

Expected:
- Agent responds with formatted loan summary
- `Generating Excel workbook...` printed
- `Saved: loan-export-2026-02-25.xlsx` printed
- File exists in current directory

**Step 3: Commit**

```bash
git add src/index.js
git commit -m "feat: CLI Excel export — detect intent, save .xlsx to cwd"
```

---

## Task 6: HTTP `POST /export` endpoint in `src/server.js`

**Files:**
- Modify: `src/server.js`

**Step 1: Add import and new route**

Find the imports at top of `src/server.js`:
```js
import { createServer } from "http";
import { FinCalAgent }  from "./agent.js";
```

Replace with:
```js
import { createServer } from "http";
import { FinCalAgent }  from "./agent.js";
import { buildWorkbook, toBuffer } from "./excel-export.js";
```

Find the 404 handler at the bottom of the request handler:
```js
  send(res, 404, { error: "Not found" });
```

Add the new route **before** that line:

```js
  if (req.method === "POST" && url.pathname === "/export") {
    try {
      const { query, session_id = "default" } = await body(req);
      if (!query) return send(res, 400, { error: "Missing query" });

      const agent    = getAgent(session_id);
      await agent.ask(query);   // populates agent.lastLoanData

      if (!agent.lastLoanData) {
        return send(res, 422, { error: "No loan data found — query must describe a loan calculation" });
      }

      const wb  = await buildWorkbook(agent.lastLoanData);
      const buf = await toBuffer(wb);
      const filename = `loan-export-${new Date().toISOString().slice(0,10)}.xlsx`;

      res.writeHead(200, {
        "Content-Type":        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename="${filename}"`,
        "Content-Length":      buf.length,
        "Access-Control-Allow-Origin": "*",
      });
      res.end(buf);
    } catch (e) {
      return send(res, 500, { error: e.message });
    }
  }
```

Also update the docs route to advertise the new endpoint. Find:
```js
        "POST /reset":     "{ session_id? } → resets conversation",
```

Add after it:
```js
        "POST /export":    "{ query, session_id? } → .xlsx file download",
```

**Step 2: Manual test**

Start the server:
```bash
ANTHROPIC_API_KEY=<your-key> node src/server.js
```

In another terminal:
```bash
curl -X POST http://localhost:3000/export \
  -H "Content-Type: application/json" \
  -d '{"query": "EMI for 10 Lakh at 15% for 5 years"}' \
  --output loan-export.xlsx
```

Expected: `loan-export.xlsx` saved locally, non-zero bytes, opens in Excel with three sheets.

**Step 3: Commit**

```bash
git add src/server.js
git commit -m "feat: add POST /export endpoint — returns xlsx binary download"
```

---

## Task 7: Extend integration test in `src/demo.js`

Add a real-data Excel export test (requires network access to FinCalEngine API).

**Files:**
- Modify: `src/demo.js`

**Step 1: Add real-data export test**

After the existing `testExcelExport()` function, add:

```js
async function testExcelExportRealData() {
  console.log(`\n${"─".repeat(60)}`);
  console.log("TEST: Excel export — real FinCalEngine data");
  console.log("─".repeat(60));

  // Get real data from API
  const bash   = createBash();
  const result = await bash.exec(
    `fincal-emi --principal 1000000 --rate 15 --term 60 --start 2026-02-25 --emi 2026-03-25`
  );
  if (result.exitCode !== 0) {
    console.log("⚠️  Skipped: FinCalEngine API unavailable");
    return;
  }

  const loanData  = JSON.parse(result.stdout.trim());
  loanData.loan   = { principal: 1000000, rate: "15", termMths: 60 };

  const wb   = await buildWorkbook(loanData);
  const path = "/tmp/fincal-realdata-test.xlsx";
  await writeToFile(wb, path);

  if (!existsSync(path))           throw new Error("File not created");
  if (statSync(path).size < 5000)  throw new Error("File suspiciously small");

  // Verify schedule row count matches loan term
  const schedSheet = wb.getWorksheet("Schedule");
  const dataRows   = schedSheet.rowCount - 1; // subtract header
  if (dataRows !== 60) throw new Error(`Expected 60 schedule rows, got ${dataRows}`);

  unlinkSync(path);
  console.log(`✅ Real-data export test passed (${dataRows} schedule rows)`);
}
```

Add `await testExcelExportRealData();` in `main()` after `testExcelExport()`.

**Step 2: Run full test suite**

```bash
node src/demo.js
```

Expected: all tests pass including the two Excel export tests.

**Step 3: Commit**

```bash
git add src/demo.js
git commit -m "test: add real-data Excel export integration test to demo.js"
```

---

## Done

At this point:
- `node src/index.js` → type any loan query + "export to excel" → `.xlsx` saved to cwd
- `node src/server.js` → `POST /export` returns binary xlsx download
- `node src/demo.js` → all tests pass including Excel export
- Workbook has Summary + Schedule + Charts sheets with embedded stacked bar and line charts
