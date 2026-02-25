// src/index.js
// Interactive CLI for FinCalAgent
// Run: ANTHROPIC_API_KEY=sk-... node src/index.js

import { createInterface } from "readline";
import { join }            from "path";
import { FinCalAgent }     from "./agent.js";
import { buildWorkbook, writeToFile } from "./excel-export.js";

if (!process.env.ANTHROPIC_API_KEY) {
  console.error("ERROR: Set ANTHROPIC_API_KEY environment variable");
  process.exit(1);
}

const verbose = process.argv.includes("--verbose") || process.argv.includes("-v");
const agent   = new FinCalAgent({ verbose });

const rl = createInterface({ input: process.stdin, output: process.stdout });

console.log();
console.log("╔══════════════════════════════════════════════════════════╗");
console.log("║           FinCalAgent  — Loan Calculator AI              ║");
console.log("║   just-bash sandbox  ·  FinCalEngine API  ·  Claude      ║");
console.log("╚══════════════════════════════════════════════════════════╝");
console.log();
console.log("  Commands: 'reset' — clear conversation | 'exit' — quit");
console.log("  Try:  EMI for 50 Lakh at 15% for 12 years");
console.log("  Try:  Add fees 5000 and step up 10% after year 10");
console.log("  Try:  Compare flat vs step-up structure");
console.log("  Try:  Export to Excel");
console.log();

function prompt() {
  rl.question("You: ", async (input) => {
    const q = input.trim();
    if (!q) return prompt();

    if (q.toLowerCase() === "exit") { rl.close(); return; }

    if (q.toLowerCase() === "reset") {
      agent.reset();
      console.log("[Conversation reset]\n");
      return prompt();
    }

    console.log("\nAgent: thinking...\n");

    const wantsExcel = /export|xlsx|excel/i.test(q);

    try {
      const answer = await agent.ask(q);
      console.log("Agent:", answer, "\n");

      if (wantsExcel && agent.lastLoanData) {
        const filename = `loan-export-${new Date().toISOString().slice(0, 10)}.xlsx`;
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

    prompt();
  });
}

prompt();
