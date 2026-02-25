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
  headerBg:   "FF1F2937",
  headerFont: "FFFFFFFF",
  accentBg:   "FFF0F9FF",
  posNum:     "FF166534",
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
  ws.columns = [{ width: 28 }, { width: 24 }];

  ws.mergeCells("A1:B1");
  const title = ws.getCell("A1");
  title.value = "Loan Summary";
  title.font  = { bold: true, size: 14, color: { argb: C.headerFont } };
  title.fill  = { type: "pattern", pattern: "solid", fgColor: { argb: C.headerBg } };
  title.alignment = { horizontal: "center", vertical: "middle" };
  ws.getRow(1).height = 28;

  const tot   = data.totals[0];
  const ret   = data.returns;
  const sched = data.amort_schedule.filter(r => r["Installment No"] > 0);
  const first = sched[0];
  const loan  = data.loan ?? {};

  const rows = [
    ["Principal",       loan.principal ? `₹${inr(loan.principal)}` : "—"],
    ["Interest Rate",   loan.rate      ? `${loan.rate}% p.a.`       : "—"],
    ["Term",            loan.termMths  ? `${loan.termMths} months`   : "—"],
    ["─────────────────────────", ""],
    ["First EMI",       `₹${inr(first?.["Installment Amount"] ?? 0)}`],
    ["Total Repayment", `₹${inr(tot["Installment Amount"])}`],
    ["Total Interest",  `₹${inr(tot["Interest"])}`],
    ["Total Principal", `₹${inr(tot["Principal"])}`],
    ...(tot["Fees"] > 0       ? [["Total Fees",       `₹${inr(tot["Fees"])}`]]       : []),
    ...(tot["Subvention"] > 0 ? [["Total Subvention", `₹${inr(tot["Subvention"])}`]] : []),
    ["─────────────────────────", ""],
    ["IRR",    ret["IRR"]],
    ["NDIRR",  ret["NDIRR"]],
    ["XIRR",   ret["XIRR"]],
    ["NDXIRR", ret["NDXIRR"]],
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
      [1, 2].forEach(col =>
        row.getCell(col).fill = { type: "pattern", pattern: "solid", fgColor: { argb: C.accentBg } }
      );
    }
    [1, 2].forEach(col => { row.getCell(col).border = borderAll(); });
  });
}

// ── Sheet 2: Schedule ─────────────────────────────────────────────────────────
function buildScheduleSheet(wb, data) {
  const ws = wb.addWorksheet("Schedule");

  const headers = ["Period", "Date", "Opening Balance", "EMI", "Interest", "Principal", "Closing Balance"];
  ws.columns = [
    { width: 8  }, { width: 14 }, { width: 20 },
    { width: 18 }, { width: 16 }, { width: 16 }, { width: 20 },
  ];

  const hRow = ws.getRow(1);
  headers.forEach((h, i) => {
    const cell = hRow.getCell(i + 1);
    cell.value     = h;
    cell.font      = { bold: true, color: { argb: C.headerFont } };
    cell.fill      = { type: "pattern", pattern: "solid", fgColor: { argb: C.headerBg } };
    cell.alignment = { horizontal: "center" };
    cell.border    = borderAll();
  });
  hRow.height = 20;

  const sched = data.amort_schedule.filter(r => r["Installment No"] > 0);
  sched.forEach((r, i) => {
    const row  = ws.getRow(i + 2);
    const vals = [
      r["Installment No"], r["Date"],
      r["Opening Balance"], r["Installment Amount"],
      r["Interest"], r["Principal"], r["Closing Balance"],
    ];
    vals.forEach((v, col) => {
      const cell  = row.getCell(col + 1);
      cell.value  = v;
      cell.border = borderAll();
      if (col >= 2) {
        cell.numFmt    = "#,##0.00";
        cell.alignment = { horizontal: "right" };
        cell.font      = { color: { argb: C.posNum } };
      }
    });
    if (i % 2 === 0) {
      [3, 4, 5, 6, 7].forEach(col =>
        row.getCell(col).fill = { type: "pattern", pattern: "solid", fgColor: { argb: C.accentBg } }
      );
    }
  });

  ws.autoFilter = { from: "A1", to: "G1" };
  ws.views = [{ state: "frozen", ySplit: 1 }];
}

// ── Sheet 3: Charts placeholder ───────────────────────────────────────────────
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
  wb.creator = "FinCalAgent";
  wb.created = new Date();

  buildSummarySheet(wb, loanData);
  buildScheduleSheet(wb, loanData);
  buildChartsSheet(wb);

  // Store schedule row count for reliable chart range computation
  wb._fincalScheduleRows = loanData.amort_schedule.filter(r => r["Installment No"] > 0).length;

  return wb;
}

// ── Public: writeToFile ───────────────────────────────────────────────────────
export async function writeToFile(workbook, filepath) {
  const tmp = join(tmpdir(), `fincal-${Date.now()}.xlsx`);
  await workbook.xlsx.writeFile(tmp);

  try {
    const withCharts = await injectCharts(tmp, workbook);
    writeFileSync(tmp + ".final", withCharts);
    renameSync(tmp + ".final", filepath);
  } catch (e) {
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
  const scheduleRowCount = workbook._fincalScheduleRows ?? (workbook.getWorksheet("Schedule").rowCount - 1);

  const zip = await JSZip.loadAsync(inputBuf);

  // Find which sheet index is "Charts" (1-indexed in xl/worksheets/sheetN.xml)
  let chartsSheetIdx = 0;
  workbook.worksheets.forEach((ws, i) => {
    if (ws.name === "Charts") chartsSheetIdx = i + 1;
  });
  if (!chartsSheetIdx) throw new Error("Charts worksheet not found in workbook");

  const sheetRelsPath = `xl/worksheets/_rels/sheet${chartsSheetIdx}.xml.rels`;

  // Add chart XML, drawing, and rels files
  zip.file("xl/charts/chart1.xml",  chartXmlStackedBar(scheduleRowCount));
  zip.file("xl/charts/chart2.xml",  chartXmlLineBalance(scheduleRowCount));
  zip.file("xl/drawings/drawing1.xml",       drawingXml());
  zip.file("xl/drawings/_rels/drawing1.xml.rels", drawingRelsXml());

  // Link the Charts sheet to the drawing
  zip.file(sheetRelsPath, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>`);

  // Inject <drawing> reference into the Charts sheet XML so Excel actually renders the charts.
  // Without this tag in the worksheet XML, Excel ignores the .rels relationship entirely.
  const sheetXmlPath = `xl/worksheets/sheet${chartsSheetIdx}.xml`;
  const sheetXml     = await zip.file(sheetXmlPath).async("string");
  const sheetXmlUpd  = sheetXml.replace(
    "</worksheet>",
    `<drawing r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></worksheet>`
  );
  zip.file(sheetXmlPath, sheetXmlUpd);

  // Register new parts in [Content_Types].xml
  const ctRaw     = await zip.file("[Content_Types].xml").async("string");
  const chartCt   = `<Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
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
