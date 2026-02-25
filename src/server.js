// src/server.js
// HTTP API server for FinCalAgent
// Run: ANTHROPIC_API_KEY=sk-... node src/server.js
// POST /calculate  { "query": "...", "session_id": "optional" }

import { createServer } from "http";
import { FinCalAgent }  from "./agent.js";
import { buildWorkbook, toBuffer } from "./excel-export.js";

if (!process.env.ANTHROPIC_API_KEY) {
  console.error("ERROR: Set ANTHROPIC_API_KEY environment variable");
  process.exit(1);
}

const PORT    = Number(process.env.PORT ?? 3000);
const sessions = new Map();          // session_id → FinCalAgent

function getAgent(sid) {
  if (!sessions.has(sid)) sessions.set(sid, new FinCalAgent());
  return sessions.get(sid);
}

function body(req) {
  return new Promise((res, rej) => {
    let s = "";
    req.on("data", (c) => (s += c));
    req.on("end", () => { try { res(JSON.parse(s)); } catch { rej(new Error("Bad JSON")); } });
    req.on("error", rej);
  });
}

function send(res, status, data) {
  res.writeHead(status, { "Content-Type": "application/json", "Access-Control-Allow-Origin": "*" });
  res.end(JSON.stringify(data, null, 2));
}

createServer(async (req, res) => {
  if (req.method === "OPTIONS") {
    res.writeHead(204, { "Access-Control-Allow-Origin": "*", "Access-Control-Allow-Methods": "POST,GET,OPTIONS", "Access-Control-Allow-Headers": "Content-Type" });
    return res.end();
  }

  const url = new URL(req.url, `http://localhost:${PORT}`);

  if (req.method === "GET" && url.pathname === "/health") {
    return send(res, 200, { status: "ok", service: "FinCalAgent", time: new Date().toISOString() });
  }

  if (req.method === "GET" && url.pathname === "/") {
    return send(res, 200, {
      service: "FinCalAgent HTTP API",
      endpoints: {
        "GET /health": "Health check",
        "POST /calculate": "{ query, session_id? } → { response, session_id }",
        "POST /reset":     "{ session_id? } → resets conversation",
        "POST /export":    "{ query, session_id? } → .xlsx file download",
      },
      example: {
        method: "POST", url: "/calculate",
        body:   { query: "EMI for 50 Lakh at 15% for 12 years", session_id: "u1" },
      },
    });
  }

  if (req.method === "POST" && url.pathname === "/calculate") {
    try {
      const { query, session_id = "default" } = await body(req);
      if (!query) return send(res, 400, { error: "Missing query" });
      const agent    = getAgent(session_id);
      const response = await agent.ask(query);
      return send(res, 200, { session_id, response, time: new Date().toISOString() });
    } catch (e) {
      return send(res, 500, { error: e.message });
    }
  }

  if (req.method === "POST" && url.pathname === "/reset") {
    const { session_id = "default" } = await body(req).catch(() => ({}));
    sessions.delete(session_id);
    return send(res, 200, { message: `Session '${session_id}' reset` });
  }

  if (req.method === "POST" && url.pathname === "/export") {
    try {
      const { query, session_id = "default" } = await body(req);
      if (!query) return send(res, 400, { error: "Missing query" });

      const agent = getAgent(session_id);
      await agent.ask(query);   // populates agent.lastLoanData

      if (!agent.lastLoanData) {
        return send(res, 422, { error: "No loan data found — query must describe a loan calculation" });
      }

      const wb       = await buildWorkbook(agent.lastLoanData);
      const buf      = await toBuffer(wb);
      const filename = `loan-export-${new Date().toISOString().slice(0, 10)}.xlsx`;

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

  send(res, 404, { error: "Not found" });
}).listen(PORT, () => {
  console.log();
  console.log(`╔══════════════════════════════════════════════════════════╗`);
  console.log(`║       FinCalAgent HTTP Server — port ${PORT}               ║`);
  console.log(`╚══════════════════════════════════════════════════════════╝`);
  console.log();
  console.log(`  Health : http://localhost:${PORT}/health`);
  console.log(`  Docs   : http://localhost:${PORT}/`);
  console.log();
  console.log(`  Example:`);
  console.log(`  curl -X POST http://localhost:${PORT}/calculate \\`);
  console.log(`       -H "Content-Type: application/json" \\`);
  console.log(`       -d '{"query":"EMI for 50 Lakh at 15% for 12 years"}'`);
  console.log();
});
