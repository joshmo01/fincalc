// src/agent.js
// FinCalAgent — agentic loop with Anthropic API + just-bash sandbox
//
// Architecture:
//   User query
//     → Claude (tool_use: bash)
//       → just-bash Bash instance (custom fincal-* commands wired in)
//         → FinCalEngine HTTP API (called inside custom command JS bodies)
//           → fincal-summary formats output
//             → Claude writes final answer
//
// No Claude Desktop. No MCP. Standalone Node.js — runs anywhere.

import Anthropic from "@anthropic-ai/sdk";
import { Bash }  from "just-bash";
import { fincalCommands } from "./commands.js";

const anthropic = new Anthropic({
  apiKey: process.env.ANTHROPIC_API_KEY,
});

const MODEL      = "claude-sonnet-4-5-20250929";
const MAX_TOKENS = 4096;
const MAX_ROUNDS = 8;

// ─── Anthropic tool definition ─────────────────────────────────────────────────
const BASH_TOOL = {
  name: "bash",
  description: `Execute bash scripts in a sandboxed environment.
Custom FinCalEngine commands are pre-installed:
  fincal-emi       - Standard fixed EMI loan
  fincal-stepup    - Step-up / moratorium / multi-phase loan
  fincal-balloon   - Balloon payment loan
  fincal-summary   - Format raw JSON output (always pipe into this)
  fincal-compare   - Side-by-side comparison of structures
Run: cat /docs/fincal.txt  for full syntax reference.`,
  input_schema: {
    type: "object",
    properties: {
      script: {
        type:        "string",
        description: "Bash script. Use fincal-* commands; pipe through fincal-summary.",
      },
    },
    required: ["script"],
  },
};

// ─── System prompt ─────────────────────────────────────────────────────────────
const TODAY      = new Date().toISOString().slice(0, 10);
const NEXT_MONTH = (() => {
  const d = new Date(); d.setMonth(d.getMonth() + 1);
  return d.toISOString().slice(0, 10);
})();

const SYSTEM = `You are FinCalAgent, an expert loan calculation AI for Indian financial services.

You have a bash tool with FinCalEngine commands pre-installed. Use it to compute repayment schedules.

Available commands:
  fincal-emi       Fixed EMI
  fincal-stepup    Step-up / moratorium / multi-phase
  fincal-balloon   Balloon payment
  fincal-summary   Format results (always pipe into this at the end)
  fincal-compare   Compare structures

Run "cat /docs/fincal.txt" for complete flag reference.

Workflow:
1. Parse the user request (principal, rate, term, structure, fees, etc.)
2. Pick the right command; consult /docs/fincal.txt if unsure
3. Write bash and run it — always end with | fincal-summary
4. Present formatted results with insights

Indian conventions:
- "10 Lakh" = 1000000, "1 Crore" = 10000000
- Default: monthly frequency, Straight Line, 30/360
- Default dates if not given: start=${TODAY}, emi=${NEXT_MONTH}

Critical rules:
- fincal-stepup --phases periods MUST sum exactly to --term
- --phases JSON must be valid; wrap in single quotes in bash
- Always pipe raw JSON through fincal-summary before presenting

Comparison pattern:
  FLAT=$(fincal-emi ...)
  STEP=$(fincal-stepup ...)
  printf '%s\\n%s\\n' "$FLAT" "$STEP" | fincal-compare --labels "Flat,Step-Up"`;

// ─── Bash executor ──────────────────────────────────────────────────────────────
const FINCAL_DOC = `
FINCALENGINE BASH COMMANDS
==========================
fincal-emi   --principal N --rate N --term N --start YYYY-MM-DD --emi YYYY-MM-DD
             [--fees N] [--commission N] [--subvention N] [--freq monthly|quarterly|...]

fincal-stepup  same as fincal-emi PLUS:
             --phases '[{"type":"Interest-Only","periods":12,"growth":0},
                        {"type":"Step","periods":12,"growth":0},
                        {"type":"Step","periods":60,"growth":15},
                        {"type":"Step","periods":36,"growth":20}]'
             Phase types: "Step" | "Interest-Only" | "Balloon/Known CashFlow"
             Sum of periods MUST equal --term. growth=0 for base phase.

fincal-balloon  same as fincal-emi PLUS:
             --regular-periods N --balloon-periods N --balloon-amount N
             (regular + balloon = term)

fincal-summary   pipe JSON from above commands here for display
fincal-compare   pipe 2+ newline-separated JSON results; --labels "A,B"

EXAMPLES:
  fincal-emi --principal 5000000 --rate 15 --term 144 --start 2026-02-20 --emi 2026-03-20 | fincal-summary

  fincal-stepup --principal 1000000 --rate 15 --term 120 --start 2026-02-20 --emi 2026-03-20 \\
    --phases '[{"type":"Interest-Only","periods":12,"growth":0},{"type":"Step","periods":108,"growth":0}]' \\
    | fincal-summary

  fincal-stepup --principal 1000000 --rate 15 --term 120 --start 2026-02-20 --emi 2026-03-20 \\
    --phases '[{"type":"Interest-Only","periods":12,"growth":0},{"type":"Step","periods":12,"growth":0},{"type":"Step","periods":60,"growth":15},{"type":"Step","periods":36,"growth":20}]' \\
    | fincal-summary

  FLAT=$(fincal-emi --principal 1000000 --rate 12 --term 60 --start 2026-02-20 --emi 2026-03-20)
  STEP=$(fincal-stepup --principal 1000000 --rate 12 --term 60 --start 2026-02-20 --emi 2026-03-20 \\
    --phases '[{"type":"Step","periods":24,"growth":0},{"type":"Step","periods":36,"growth":10}]')
  printf '%s\\n%s\\n' "$FLAT" "$STEP" | fincal-compare --labels "Flat EMI,Step-Up 10%"
`;

async function execBash(script, verbose = false) {
  const bash = new Bash({
    customCommands: fincalCommands,
    files:          { "/docs/fincal.txt": FINCAL_DOC },
    executionLimits: { maxCommandCount: 500, maxLoopIterations: 500 },
  });

  if (verbose) console.log(`\n[Bash script]\n${script}\n`);

  const result = await bash.exec(script);

  if (verbose) {
    if (result.stdout) console.log("[stdout]", result.stdout.slice(0, 600));
    if (result.stderr) console.log("[stderr]", result.stderr.slice(0, 300));
    console.log("[exit]", result.exitCode);
  }

  return result;
}

// ─── Agent class ────────────────────────────────────────────────────────────────
export class FinCalAgent {
  constructor({ verbose = false } = {}) {
    this.verbose      = verbose;
    this.history      = [];
    this.lastLoanData = null;   // populated after each successful FinCalEngine call
  }

  /** Send a query; returns final text response */
  async ask(userMessage) {
    this.history.push({ role: "user", content: userMessage });

    for (let round = 0; round < MAX_ROUNDS; round++) {
      if (this.verbose) console.log(`\n[Agent] Round ${round + 1}`);

      const response = await anthropic.messages.create({
        model:      MODEL,
        max_tokens: MAX_TOKENS,
        system:     SYSTEM,
        tools:      [BASH_TOOL],
        messages:   this.history,
      });

      if (this.verbose) console.log(`[Agent] stop_reason=${response.stop_reason}`);

      this.history.push({ role: "assistant", content: response.content });

      if (response.stop_reason === "tool_use") {
        const toolResults = [];

        for (const block of response.content) {
          if (block.type !== "tool_use") continue;

          const script = block.input?.script ?? "";
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
            type:        "tool_result",
            tool_use_id: block.id,
            content:     JSON.stringify({ stdout, stderr, exitCode }),
          });
        }

        this.history.push({ role: "user", content: toolResults });
        continue;
      }

      if (response.stop_reason === "end_turn") {
        return response.content
          .filter((b) => b.type === "text")
          .map((b) => b.text)
          .join("\n");
      }
    }

    return "[Agent: max rounds reached — please try again]";
  }

  reset() { this.history = []; this.lastLoanData = null; }
}
