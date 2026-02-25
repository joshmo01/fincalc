// src/demo.js
// Tests the just-bash sandbox + FinCalEngine commands directly,
// without calling the Anthropic API. Use this to verify the plumbing works.
//
// Run: node src/demo.js

import { Bash } from "just-bash";
import { fincalCommands } from "./commands.js";
import { buildWorkbook, writeToFile } from "./excel-export.js";
import { existsSync, statSync, unlinkSync } from "fs";

const FINCAL_DOC = "See commands.js for reference";

function createBash() {
  return new Bash({
    customCommands: fincalCommands,
    files: { "/docs/fincal.txt": FINCAL_DOC },
  });
}

async function run(label, script) {
  console.log(`\n${"─".repeat(60)}`);
  console.log(`TEST: ${label}`);
  console.log("─".repeat(60));

  const bash = createBash();
  const r    = await bash.exec(script);

  if (r.exitCode === 0) {
    console.log(r.stdout);
  } else {
    console.error("FAILED:", r.stderr);
  }
}

async function testExcelExport() {
  console.log(`\n${"─".repeat(60)}`);
  console.log("TEST: Excel export — Summary + Schedule sheets");
  console.log("─".repeat(60));

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

  if (!existsSync(path))          throw new Error("File not created");
  if (statSync(path).size < 1000) throw new Error("File too small — likely empty");

  const sheetNames = wb.worksheets.map(s => s.name);
  if (!sheetNames.includes("Summary"))  throw new Error("Missing Summary sheet");
  if (!sheetNames.includes("Schedule")) throw new Error("Missing Schedule sheet");
  if (!sheetNames.includes("Charts"))   throw new Error("Missing Charts sheet");

  unlinkSync(path);
  console.log("✅ Excel export test passed");
}

async function main() {
  console.log("FinCalAgent — Sandbox Integration Tests");
  console.log("(No Anthropic API key required for this test)\n");
  console.log("⚠️  Requires network access to http://13.127.48.98:8000");

  await run(
    "Simple flat EMI — 50L @ 15% / 12 years",
    `fincal-emi --principal 5000000 --rate 15 --term 144 \
--start 2026-02-20 --emi 2026-03-20 | fincal-summary`
  );

  await run(
    "Interest moratorium 1yr + EMI 9yr — 10L @ 15%",
    `fincal-stepup --principal 1000000 --rate 15 --term 120 \
--start 2026-02-20 --emi 2026-03-20 \
--phases '[{"type":"Interest-Only","periods":12,"growth":0},{"type":"Step","periods":108,"growth":0}]' \
| fincal-summary`
  );

  await run(
    "Step-up: moratorium + base + 15% + 20% — 10L @ 15% / 10 years",
    `fincal-stepup --principal 1000000 --rate 15 --term 120 \
--start 2026-02-20 --emi 2026-03-20 \
--phases '[{"type":"Interest-Only","periods":12,"growth":0},{"type":"Step","periods":12,"growth":0},{"type":"Step","periods":60,"growth":15},{"type":"Step","periods":36,"growth":20}]' \
| fincal-summary`
  );

  await run(
    "Balloon payment — 20L @ 12% / 5 years, balloon 5L at end",
    `fincal-balloon --principal 2000000 --rate 12 --term 60 \
--start 2026-02-20 --emi 2026-03-20 \
--regular-periods 59 --balloon-periods 1 --balloon-amount 500000 \
| fincal-summary`
  );

  await run(
    "Comparison — Flat EMI vs Step-Up 10% — 10L @ 12% / 5 years",
    `FLAT=$(fincal-emi --principal 1000000 --rate 12 --term 60 \
--start 2026-02-20 --emi 2026-03-20)
STEP=$(fincal-stepup --principal 1000000 --rate 12 --term 60 \
--start 2026-02-20 --emi 2026-03-20 \
--phases '[{"type":"Step","periods":24,"growth":0},{"type":"Step","periods":36,"growth":10}]')
printf '%s\n%s\n' "$FLAT" "$STEP" | fincal-compare --labels "Flat EMI,Step-Up 10%"`
  );

  await testExcelExport();

  console.log("\n✅ All tests complete.");
}

main().catch((e) => { console.error("Demo error:", e); process.exit(1); });
