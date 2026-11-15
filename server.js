// Local HTTPS server:
//  - task pane for the Word add-in (Office requires HTTPS, even on localhost)
//  - /api/review: playbook review proxy (Claude key never reaches the browser)
//  - /api/intake + Gmail poller: emails -> matters -> versions -> analysis
//  - /api/matters, /api/triage: version registry and the "needs a matter" queue

import express from "express";
import https from "node:https";
import http from "node:http";
import { readFileSync, readdirSync } from "node:fs";
import { fileURLToPath } from "node:url";
import path from "node:path";
import { reviewUnits, loadPlaybook, liveMode, MODEL } from "./server/review-service.js";
import { createContext } from "./server/context.js";
import { processMessage } from "./server/intake.js";
import { startGmailPoller } from "./server/gmail-poller.js";

const root = path.dirname(fileURLToPath(import.meta.url));
const PORT = Number(process.env.PORT || 3000);
const ctx = createContext();

const app = express();
app.use(express.json({ limit: "2mb" }));
app.use(express.static(path.join(root, "public")));
app.use("/vendor/jszip", express.static(path.join(root, "node_modules/jszip/dist")));
app.use("/playbooks", express.static(path.join(root, "playbooks")));
app.use("/samples", express.static(path.join(root, "samples"), { index: false }));

// Browser preview without Word (PREVIEW=1): task pane + Office.js stub + sample doc
if (process.env.PREVIEW) {
  app.get("/preview.html", (_req, res) => {
    const html = readFileSync(path.join(root, "public/taskpane.html"), "utf8").replace(
      /<script src="https:\/\/appsforoffice[^"]+"><\/script>\s*<script src="\/vendor\/jszip\/jszip.min.js"><\/script>/,
      '<script src="/vendor/jszip/jszip.min.js"></script><script src="/dev/office-stub.js"></script>',
    );
    res.type("html").send(html);
  });
  app.use("/dev", express.static(path.join(root, "tests")));
}

const wrap = (fn) => async (req, res) => {
  try { await fn(req, res); } catch (err) { console.error(err); res.status(502).json({ error: String(err.message || err) }); }
};

// ---- playbook review (used by the task pane) --------------------------------

app.get("/api/playbooks", (_req, res) => {
  const list = readdirSync(path.join(root, "playbooks"))
    .filter((f) => f.endsWith(".json"))
    .map((f) => {
      const p = loadPlaybook(f.replace(/\.json$/, ""));
      return { id: p.id, contractType: p.contractType, version: p.version, rules: p.rules.length };
    });
  res.json({ playbooks: list, mode: liveMode() ? "live" : "mock", model: MODEL });
});

app.post("/api/review", wrap(async (req, res) => {
  const started = Date.now();
  const { playbookId, units, mock } = req.body || {};
  if (!Array.isArray(units) || !units.length) return res.status(400).json({ error: "no review units" });
  if (units.length > 200) return res.status(400).json({ error: "too many changes in one request (max 200)" });
  const playbook = loadPlaybook(playbookId || "mutual-nda");
  const result = await reviewUnits(playbook, units, { mock });
  // audit line: what was decided, never the contract text
  console.log(JSON.stringify({
    at: new Date().toISOString(), mode: result.mode, model: result.model, playbook: `${playbook.id}@${playbook.version}`,
    units: units.length, ms: Date.now() - started, harnessIssues: result.issues.length,
    verdicts: result.findings.reduce((a, f) => ((a[f.verdict] = (a[f.verdict] || 0) + 1), a), {}),
  }));
  res.json(result);
}));

// ---- email intake ------------------------------------------------------------

app.post("/api/intake", express.raw({ type: () => true, limit: "30mb" }), wrap(async (req, res) => {
  if (!req.body?.length) return res.status(400).json({ error: "send the raw email (RFC 822) as the request body" });
  res.json(await processMessage(req.body, ctx));
}));

app.get("/api/matters", (_req, res) => {
  res.json(ctx.store.db().matters.map((m) => ({
    ...m,
    versions: ctx.store.versions(m.id).length,
    current: ctx.store.versions(m.id).filter((v) => v.status === "active").map((v) => `${v.label} (${v.side})`),
  })));
});

app.get("/api/matters/:id", (req, res) => {
  const m = ctx.store.matter(req.params.id);
  if (!m) return res.status(404).json({ error: "unknown matter" });
  const strip = ({ afterTexts, beforeTexts, ...v }) => v; // keep payload small
  res.json({ ...m, versions: ctx.store.versions(m.id).map(strip) });
});

app.get("/api/versions/:id/file", (req, res) => {
  const v = ctx.store.version(req.params.id);
  if (!v) return res.status(404).end();
  res.download(ctx.store.blobPath("files", v.fileHash, ".docx"), v.fileName);
});

app.get("/api/triage", (_req, res) => res.json(ctx.store.openTriage()));

app.post("/api/triage/:id/assign", wrap(async (req, res) => {
  const t = ctx.store.triage(req.params.id);
  if (!t || t.status !== "open") return res.status(404).json({ error: "no open triage item" });
  const matterId = req.body?.matterId;
  if (!ctx.store.matter(matterId)) return res.status(400).json({ error: "unknown matterId" });
  const report = await processMessage(ctx.store.readBlob("mail", t.emlHash, ".eml"), { ...ctx, forceMatterId: matterId });
  ctx.store.closeTriage(t.id, { assignedTo: matterId, closedAt: new Date().toISOString() });
  res.json(report);
}));

// ---- start --------------------------------------------------------------------

async function start() {
  const banner = `[${liveMode() ? "live: " + MODEL : "mock reviewer - set ANTHROPIC_API_KEY for live review"}] config: ${path.relative(root, ctx.config._file)}`;
  if (process.env.HTTP_ONLY) {
    http.createServer(app).listen(PORT, () => console.log(`redliningPal on http://localhost:${PORT} ${banner}`));
  } else {
    const certs = await import("office-addin-dev-certs");
    const options = await certs.getHttpsServerOptions();
    https.createServer(options, app).listen(PORT, () => console.log(`redliningPal on https://localhost:${PORT} ${banner}`));
  }
  startGmailPoller(ctx);
}

start();
