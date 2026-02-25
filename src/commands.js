// src/commands.js
// just-bash custom commands that call FinCalEngine API internally.
//
// Key insight: defineCommand lets us add TypeScript/JS functions as
// first-class bash commands. The command body runs in the Node.js
// host environment, so it can do real HTTP fetches — while the
// bash script layer stays fully sandboxed.
//
// The AI agent only needs to write bash scripts using these commands.
// It never needs to know about the API details.

import { defineCommand } from "just-bash";
import { calcRepayments, buildPayload } from "./fincal-api.js";

// Module-level storage for the last successful FinCalEngine result.
// agent.js reads this after each bash execution to capture data for Excel export.
let _lastResult = null;
export function getLastResult() { return _lastResult; }

// ─── helpers ──────────────────────────────────────────────────────────────────

function ok(data) {
  return { stdout: JSON.stringify(data, null, 2) + "\n", stderr: "", exitCode: 0 };
}

function err(msg) {
  return { stdout: "", stderr: `fincal error: ${msg}\n`, exitCode: 1 };
}

function parseArg(args, flag) {
  const i = args.indexOf(flag);
  return i !== -1 ? args[i + 1] : null;
}

function parseFloat2(s) {
  const v = parseFloat(s);
  if (isNaN(v)) throw new Error(`Expected number, got: ${s}`);
  return v;
}

function parseInt2(s) {
  const v = parseInt(s, 10);
  if (isNaN(v)) throw new Error(`Expected integer, got: ${s}`);
  return v;
}

// ─── fincal-emi ───────────────────────────────────────────────────────────────
// Standard fixed EMI loan
// Usage: fincal-emi --principal 1000000 --rate 15 --term 120 --start 2026-02-20 --emi 2026-03-20
//        [--fees 5000] [--commission 0] [--subvention 10000] [--freq monthly]

export const cmdEmi = defineCommand("fincal-emi", async (args) => {
  try {
    const principal  = parseFloat2(parseArg(args, "--principal"));
    const rate       = parseArg(args, "--rate");
    const termMths   = parseInt2(parseArg(args, "--term"));
    const startDate  = parseArg(args, "--start");
    const emiDate    = parseArg(args, "--emi");
    const fees       = parseFloat(parseArg(args, "--fees")       ?? "0");
    const commission = parseFloat(parseArg(args, "--commission") ?? "0");
    const subvention = parseFloat(parseArg(args, "--subvention") ?? "0");
    const frequency  = parseArg(args, "--freq") ?? "monthly";

    const payload = buildPayload({
      principal, rate, termMths, frequency, startDate, emiDate,
      fees, commission, subvention,
      repayStructure: [
        { "Step/Balloon/CashFlow": "Step", "No of Periods": termMths, GrwthRate: 0, Value: 0 },
      ],
    });

    const result = await calcRepayments(payload);
    _lastResult = result;
    return ok(result);
  } catch (e) {
    return err(e.message);
  }
});

// ─── fincal-stepup ────────────────────────────────────────────────────────────
// Step-up / moratorium / mixed phase loan
// Phases are passed as JSON via --phases '[...]'
// Each phase: {"type":"Step|Interest-Only|Balloon/Known CashFlow","periods":N,"growth":N,"value":N}

export const cmdStepup = defineCommand("fincal-stepup", async (args) => {
  try {
    const principal  = parseFloat2(parseArg(args, "--principal"));
    const rate       = parseArg(args, "--rate");
    const termMths   = parseInt2(parseArg(args, "--term"));
    const startDate  = parseArg(args, "--start");
    const emiDate    = parseArg(args, "--emi");
    const fees       = parseFloat(parseArg(args, "--fees")       ?? "0");
    const commission = parseFloat(parseArg(args, "--commission") ?? "0");
    const subvention = parseFloat(parseArg(args, "--subvention") ?? "0");
    const frequency  = parseArg(args, "--freq") ?? "monthly";
    const phasesRaw  = parseArg(args, "--phases");

    if (!phasesRaw) return err("--phases JSON required");

    const phases = JSON.parse(phasesRaw);

    // Validate period sum
    const totalPeriods = phases.reduce((s, p) => s + p.periods, 0);
    if (totalPeriods !== termMths) {
      return err(`Phase periods sum (${totalPeriods}) ≠ term months (${termMths})`);
    }

    const repayStructure = phases.map((p) => ({
      "Step/Balloon/CashFlow": p.type,
      "No of Periods": p.periods,
      GrwthRate: p.growth ?? 0,
      Value: p.value ?? 0,
    }));

    const payload = buildPayload({
      principal, rate, termMths, frequency, startDate, emiDate,
      fees, commission, subvention, repayStructure,
    });

    const result = await calcRepayments(payload);
    _lastResult = result;
    return ok(result);
  } catch (e) {
    return err(e.message);
  }
});

// ─── fincal-balloon ───────────────────────────────────────────────────────────
// Balloon payment loan
// Usage: fincal-balloon --principal P --rate R --term T --start D --emi D
//                       --regular-periods N --balloon-periods N --balloon-amount A

export const cmdBalloon = defineCommand("fincal-balloon", async (args) => {
  try {
    const principal      = parseFloat2(parseArg(args, "--principal"));
    const rate           = parseArg(args, "--rate");
    const termMths       = parseInt2(parseArg(args, "--term"));
    const startDate      = parseArg(args, "--start");
    const emiDate        = parseArg(args, "--emi");
    const regularPeriods = parseInt2(parseArg(args, "--regular-periods"));
    const balloonPeriods = parseInt2(parseArg(args, "--balloon-periods"));
    const balloonAmount  = parseFloat2(parseArg(args, "--balloon-amount"));
    const fees           = parseFloat(parseArg(args, "--fees")       ?? "0");
    const commission     = parseFloat(parseArg(args, "--commission") ?? "0");
    const subvention     = parseFloat(parseArg(args, "--subvention") ?? "0");

    if (regularPeriods + balloonPeriods !== termMths) {
      return err(`regular-periods + balloon-periods must equal term`);
    }

    const repayStructure = [
      { "Step/Balloon/CashFlow": "Step",                   "No of Periods": regularPeriods, GrwthRate: 0, Value: 0 },
      { "Step/Balloon/CashFlow": "Balloon/Known CashFlow", "No of Periods": balloonPeriods, GrwthRate: 0, Value: balloonAmount },
    ];

    const payload = buildPayload({
      principal, rate, termMths, startDate, emiDate,
      fees, commission, subvention, repayStructure,
    });

    const result = await calcRepayments(payload);
    _lastResult = result;
    return ok(result);
  } catch (e) {
    return err(e.message);
  }
});

// ─── fincal-summary ───────────────────────────────────────────────────────────
// Pipe raw calcRepayments JSON through this to get a human-readable summary
// Usage: fincal-emi ... | fincal-summary

export const cmdSummary = defineCommand("fincal-summary", async (_args, ctx) => {
  try {
    const raw  = ctx.stdin.trim();
    if (!raw) return err("No input — pipe fincal-emi or fincal-stepup output into fincal-summary");

    const data = JSON.parse(raw);
    const sched  = data.amort_schedule;
    const totals = data.totals[0];
    const ret    = data.returns;
    const INR    = (n) => `₹${Number(n).toLocaleString("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;

    // First non-disbursement row
    const firstEmi = sched.find((r) => r["Installment No"] > 0);

    const phases = data.summary
      .filter((p) => p["Step-Type"] !== "Disbursement")
      .map((p) => `  ${p["Step-Type"].padEnd(22)} | ${p["Date From"]} → ${p["Date To"]}  | EMI ${INR(p["Installment Amount"])}`)
      .join("\n");

    // First 3 + last 2 rows of schedule
    const rows = sched.filter((r) => r["Installment No"] > 0);
    const preview = [...rows.slice(0, 3), null, ...rows.slice(-2)];
    const schedLines = preview.map((r) => {
      if (!r) return "  ...";
      return `  #${String(r["Installment No"]).padStart(3)}  ${r["Date"]}  Opening ${INR(r["Opening Balance"])}  EMI ${INR(r["Installment Amount"])}  Interest ${INR(r["Interest"])}  Principal ${INR(r["Principal"])}`;
    }).join("\n");

    const out = [
      "═══════════════════════════════════════════════════════════",
      " KEY METRICS",
      "═══════════════════════════════════════════════════════════",
      `  First EMI        : ${INR(firstEmi?.["Installment Amount"] ?? 0)}`,
      `  Total Repayment  : ${INR(totals["Installment Amount"])}`,
      `  Total Interest   : ${INR(totals["Interest"])}`,
      `  Total Principal  : ${INR(totals["Principal"])}`,
      totals["Fees"] > 0 ? `  Total Fees        : ${INR(totals["Fees"])}` : null,
      totals["Subvention"] > 0 ? `  Total Subvention  : ${INR(totals["Subvention"])}` : null,
      `  IRR              : ${ret["IRR"]}`,
      `  NDIRR            : ${ret["NDIRR"]}`,
      `  XIRR             : ${ret["XIRR"]}`,
      `  NDXIRR           : ${ret["NDXIRR"]}`,
      "",
      "───────────────────────────────────────────────────────────",
      " PHASE SUMMARY",
      "───────────────────────────────────────────────────────────",
      phases,
      "",
      "───────────────────────────────────────────────────────────",
      " SCHEDULE PREVIEW",
      "───────────────────────────────────────────────────────────",
      schedLines,
      "═══════════════════════════════════════════════════════════",
    ].filter((l) => l !== null).join("\n");

    return { stdout: out + "\n", stderr: "", exitCode: 0 };
  } catch (e) {
    return err(e.message);
  }
});

// ─── fincal-compare ───────────────────────────────────────────────────────────
// Compare N structures. Reads newline-delimited JSON blobs from stdin
// or takes two named JSON files as args.
// Usage: { echo RESULT_A; echo RESULT_B; } | fincal-compare --labels "Flat,StepUp"

export const cmdCompare = defineCommand("fincal-compare", async (args, ctx) => {
  try {
    const labelsArg = parseArg(args, "--labels") ?? "";
    const labels    = labelsArg ? labelsArg.split(",") : [];

    const lines = ctx.stdin.trim().split(/\n(?=\{)/); // split on newline before {
    if (lines.length < 2) return err("Pipe at least two JSON results (newline-delimited)");

    const INR = (n) => `₹${Number(n).toLocaleString("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;

    const cols = lines.map((line, i) => {
      const data  = JSON.parse(line.trim());
      const tot   = data.totals[0];
      const ret   = data.returns;
      const firstEmi = data.amort_schedule.find((r) => r["Installment No"] > 0);
      return {
        label:       labels[i] ?? `Structure ${i + 1}`,
        emi:         INR(firstEmi?.["Installment Amount"] ?? 0),
        totalRepay:  INR(tot["Installment Amount"]),
        totalInt:    INR(tot["Interest"]),
        irr:         ret["IRR"],
        xirr:        ret["XIRR"],
      };
    });

    const pad = (s, n) => String(s).padEnd(n);
    const W   = 22;

    const header  = ["METRIC", ...cols.map((c) => c.label)].map((v, i) => pad(v, i === 0 ? 20 : W)).join(" | ");
    const divider = "─".repeat(header.length);

    const row = (label, key) =>
      [label, ...cols.map((c) => c[key])].map((v, i) => pad(v, i === 0 ? 20 : W)).join(" | ");

    const out = [
      "═".repeat(header.length),
      " LOAN STRUCTURE COMPARISON",
      "═".repeat(header.length),
      header,
      divider,
      row("First EMI",        "emi"),
      row("Total Repayment",  "totalRepay"),
      row("Total Interest",   "totalInt"),
      row("IRR",              "irr"),
      row("XIRR",             "xirr"),
      "═".repeat(header.length),
    ].join("\n");

    return { stdout: out + "\n", stderr: "", exitCode: 0 };
  } catch (e) {
    return err(e.message);
  }
});

// Export all commands as an array for easy consumption
export const fincalCommands = [cmdEmi, cmdStepup, cmdBalloon, cmdSummary, cmdCompare];
