# Excel Export — Design Doc

**Date:** 2026-02-25
**Status:** Approved

---

## Goal

Add Excel workbook export to FinCalAgent. A user can request an export from the CLI (file saved to disk) or via the HTTP server (binary download). The workbook contains three sheets: Summary, Repayment Schedule, and Charts.

---

## Architecture

Excel export lives in a new module `src/excel-export.js`. It sits outside the agent/sandbox loop — it receives the raw FinCalEngine JSON and produces a `.xlsx` file.

```
FinCalEngine JSON
      │
      ▼
┌─────────────────────────┐
│   excel-export.js       │
│   buildWorkbook(data)   │  ← uses exceljs
│                         │
│  Sheet 1: Summary       │  key metrics table
│  Sheet 2: Schedule      │  full repayment rows
│  Sheet 3: Charts        │  chart XML injected post-write
└──────────┬──────────────┘
           │ writes .xlsx buffer or file
           ▼
    ┌──────┴──────┐
    │             │
  CLI           HTTP
 saves file   streams buffer
 to disk      as download
```

The agent internals are unchanged. Export is triggered by the CLI/server layer after receiving loan JSON.

---

## Components

### `src/excel-export.js`

Three exported functions:

- **`buildWorkbook(loanData)`** — takes parsed FinCalEngine JSON, returns an ExcelJS `Workbook` with all three sheets populated
- **`writeToFile(workbook, filepath)`** — writes workbook to disk, then injects chart XML into the zip archive via `jszip`
- **`toBuffer(workbook)`** — same as above, returns a `Buffer` for HTTP streaming

**Chart XML injection:** ExcelJS writes a valid `.xlsx` (a zip). After writing, `jszip` reopens it and adds:
- `xl/charts/chart1.xml` — stacked bar chart (Principal vs Interest per period)
- `xl/charts/chart2.xml` — line chart (Closing Balance over time)
- Updates `xl/worksheets/_rels/sheet3.xml.rels` and `[Content_Types].xml` to register charts

### `src/commands.js`

Add `fincal-excel` sandbox command. Outputs raw FinCalEngine JSON with a sentinel marker so the CLI/server layer knows to trigger Excel export instead of text formatting.

### `src/index.js`

Detect "export" intent in the agent response → call `writeToFile()` → print saved filename.

### `src/server.js`

Add `POST /export` route: same `{ query, session_id }` body as `/calculate`, but responds with the `.xlsx` binary and appropriate headers:
```
Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
Content-Disposition: attachment; filename="loan-export-<timestamp>.xlsx"
```

### New npm dependencies

- `exceljs` — workbook/sheet/cell formatting
- `jszip` — zip manipulation for chart XML injection

---

## Data Flow

### Sheet 1 — Summary

Key metrics from FinCalEngine response top-level fields:

| Label | Example |
|---|---|
| Principal | ₹50,00,000 |
| Interest Rate | 15.00% p.a. |
| Term | 144 months |
| First EMI | ₹74,405.10 |
| Total Repayment | ₹1,07,14,334 |
| Total Interest | ₹57,14,334 |
| IRR | 15.00% |
| XIRR | 16.08% |

### Sheet 2 — Repayment Schedule

One row per period from the `repayments` array:

`Period | Date | Opening Balance | EMI | Interest | Principal | Closing Balance`

### Sheet 3 — Charts

Two charts referencing Sheet 2 data ranges directly (no data duplication):

- **Chart 1 (stacked bar):** X = Period, Series 1 = Interest, Series 2 = Principal
- **Chart 2 (line):** X = Period, Y = Closing Balance

### CLI export flow

```
User: "EMI for 50L at 15% 12 years, export to excel"
  → Agent runs fincal-emi, gets JSON
  → CLI detects export intent
  → buildWorkbook(json) → writeToFile() → inject chart XML
  → "Saved: loan-export-2026-02-25.xlsx"
```

### HTTP export flow

```
POST /export { query: "...", session_id: "u1" }
  → Agent run, JSON captured
  → toBuffer() → inject chart XML into buffer
  → Binary .xlsx response with filename header
```

---

## Error Handling

| Failure | Handling |
|---|---|
| FinCalEngine API failure | Already handled by agent; export never called |
| Missing fields in JSON | `buildWorkbook()` validates before touching ExcelJS; descriptive error thrown |
| File write failure (CLI) | Write to temp path first, rename on success; clean message to user |
| Chart XML injection failure | Graceful degradation — workbook saved without charts, console warning printed |

---

## Testing

- **Integration test (`src/demo.js`):** extend existing test — after a successful FinCalEngine call, run `buildWorkbook()` and assert the output file exists and is non-zero bytes
- **Manual smoke test:** open generated `.xlsx` in Excel/Numbers/LibreOffice, verify three sheets and chart rendering

No new test framework introduced (project has none).
