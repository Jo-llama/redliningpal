// Express app, without listening: server.js starts it, tests drive it over HTTP.
//  - /api/review: playbook review proxy (Claude key never reaches the browser)
//  - /api/intake: emails -> matters -> versions -> analysis
//  - /api/matters, /api/triage: version registry and the "needs a matter" queue
//  - /api/documents, /api/versions/:id/*: what is open in Word, its timeline,
//    and lawyer-confirmed registration of unknown documents

import express from "express";
import { readFileSync, readdirSync } from "node:fs";
import crypto from "node:crypto";
import path from "node:path";
import { reviewUnits, loadPlaybook, liveMode, MODEL } from "./review-service.js";
import { processMessage, registerDocument } from "./intake.js";
import { readDocx } from "./docx.js";
import { rankBases } from "./matching.js";
import { identifyDocument, buildTimeline } from "./timeline.js";
import { norm } from "../public/lib/ooxml.js";

const MAX_PARAGRAPHS = 20000;

/** Audit lines carry ids, counts and verdicts only - never contract text. */
const defaultAudit = (event) => console.log(JSON.stringify({ at: new Date().toISOString(), ...event }));

const bad = (message) => Object.assign(new Error(message), { status: 400 });

/**
 * @param ctx   { store, config, mock?, notify? } from createContext()
 * @param opts  { root, preview?, audit? }
 */
export function createApp(ctx, { root, preview = !!process.env.PREVIEW, audit = defaultAudit } = {}) {
  const app = express();
  app.use(express.json({ limit: "5mb" }));
  app.use(express.static(path.join(root, "public")));
  app.use("/vendor/jszip", express.static(path.join(root, "node_modules/jszip/dist")));
  app.use("/playbooks", express.static(path.join(root, "playbooks")));
  app.use("/samples", express.static(path.join(root, "samples"), { index: false }));

  // Browser preview without Word (PREVIEW=1): task pane + Office.js stub + sample doc
  if (preview) {
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
    try { await fn(req, res); } catch (err) {
      if (!err.status) console.error(err);
      res.status(err.status || 502).json({ error: String(err.message || err) });
    }
  };
  const store = ctx.store;
  const versionOr404 = (req, res) => {
    const v = store.version(req.params.id);
    if (!v) res.status(404).json({ error: "unknown version" });
    return v;
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
    const result = await reviewUnits(playbook, units, { mock: mock ?? ctx.mock });
    audit({
      event: "review", mode: result.mode, model: result.model, playbook: `${playbook.id}@${playbook.version}`,
      units: units.length, ms: Date.now() - started, harnessIssues: result.issues.length,
      verdicts: result.findings.reduce((a, f) => ((a[f.verdict] = (a[f.verdict] || 0) + 1), a), {}),
    });
    res.json(result);
  }));

  // ---- email intake ------------------------------------------------------------

  app.post("/api/intake", express.raw({ type: () => true, limit: "30mb" }), wrap(async (req, res) => {
    if (!req.body?.length) return res.status(400).json({ error: "send the raw email (RFC 822) as the request body" });
    res.json(await processMessage(req.body, ctx));
  }));

  app.get("/api/matters", (_req, res) => {
    res.json(store.db().matters.map((m) => ({
      ...m,
      versions: store.versions(m.id).length,
      current: store.versions(m.id).filter((v) => v.status === "active").map((v) => `${v.label} (${v.side})`),
    })));
  });

  app.get("/api/matters/:id", (req, res) => {
    const m = store.matter(req.params.id);
    if (!m) return res.status(404).json({ error: "unknown matter" });
    const strip = ({ afterTexts, beforeTexts, ...v }) => v; // keep payload small
    res.json({ ...m, versions: store.versions(m.id).map(strip) });
  });

  app.get("/api/versions/:id/file", (req, res) => {
    const v = versionOr404(req, res);
    if (!v) return;
    res.download(store.blobPath("files", v.fileHash, ".docx"), v.fileName);
  });

  // ---- the document open in Word ------------------------------------------------

  // body: { afterTexts: string[], beforeTexts: string[], stamp?: string } from the task pane
  app.post("/api/documents/identify", wrap(async (req, res) => {
    const { afterTexts, beforeTexts, stamp } = req.body || {};
    const texts = (a) => {
      if (!Array.isArray(a) || a.length > MAX_PARAGRAPHS || a.some((s) => typeof s !== "string")) {
        throw bad(`afterTexts and beforeTexts must be arrays of strings (max ${MAX_PARAGRAPHS})`);
      }
      return a.map(norm);
    };
    const doc = { afterTexts: texts(afterTexts), beforeTexts: texts(beforeTexts), stamp: typeof stamp === "string" ? stamp : null };
    if (doc.afterTexts.length !== doc.beforeTexts.length) throw bad("afterTexts and beforeTexts differ in length");
    const all = store.versions();
    const result = identifyDocument(doc, all);
    audit({ event: "identify", status: result.status, versionId: result.versionId, score: result.score, paragraphs: doc.afterTexts.length });
    res.json({ ...result, timeline: result.versionId ? buildTimeline(all, result.versionId) : null });
  }));

  app.get("/api/versions/:id/timeline", (req, res) => {
    if (!versionOr404(req, res)) return;
    res.json(buildTimeline(store.versions(), req.params.id));
  });

  // the changes made without track changes, as review cards (contract text: never logged)
  app.get("/api/versions/:id/untracked", (req, res) => {
    const v = versionOr404(req, res);
    if (!v) return;
    const units = (v.analysis?.units || []).filter((u) => u.source === "untracked")
      .map(({ id, section, kind, paragraphIndex, before, after, finding }) => ({ id, section, kind, paragraphIndex, before, after, finding }));
    res.json({ versionId: v.id, checked: !!v.analysis?.integrity?.checked, units });
  });

  app.post("/api/versions/:id/ack-untracked", (req, res) => {
    const v = versionOr404(req, res);
    if (!v) return;
    const count = v.analysis?.integrity?.untracked || 0;
    if (!count) return res.status(409).json({ error: "this version has no untracked changes" });
    const at = new Date().toISOString();
    store.updateVersion(v.id, { untrackedAck: { at, count } });
    audit({ event: "ack_untracked", versionId: v.id, matterId: v.matterId, untracked: count });
    res.json({ versionId: v.id, acknowledgedAt: at });
  });

  // unknown document -> triage item with the file; nothing is registered until a lawyer assigns it
  app.post("/api/documents/triage", express.raw({ type: () => true, limit: "30mb" }), wrap(async (req, res) => {
    if (!req.body?.length) throw bad("send the .docx as the request body");
    let doc;
    try { doc = await readDocx(req.body); } catch { throw bad("not a readable .docx"); }
    const all = store.versions();
    const id = identifyDocument(doc, all);
    if (id.status === "exact") {
      return res.status(409).json({ error: "already registered", versionId: id.versionId, timeline: buildTimeline(all, id.versionId) });
    }
    const fileName = String(req.get("x-file-name") || "from-word.docx").replace(/[^\w .()-]/g, "_").slice(0, 120);
    const t = store.addTriage({
      id: `T-${crypto.randomUUID().slice(0, 8)}`, kind: "document", status: "open",
      fileHash: store.putBlob("files", req.body, ".docx"), fileName, subject: fileName, from: "Word add-in",
      candidates: id.candidates, createdAt: new Date().toISOString(),
    });
    audit({ event: "triage_document", triageId: t.id, candidates: id.candidates.length, bestScore: id.score });
    res.status(201).json({ triageId: t.id, candidates: t.candidates, matters: store.db().matters.map(({ id: mid, title }) => ({ id: mid, title })) });
  }));

  // suggested bases for a document triage item, so the lawyer confirms rather than guesses
  app.get("/api/triage/:id/bases", wrap(async (req, res) => {
    const t = store.triage(req.params.id);
    if (!t || t.kind !== "document" || t.status !== "open") return res.status(404).json({ error: "no open document triage item" });
    const { matterId, side } = req.query;
    if (!store.matter(matterId)) throw bad("unknown matterId");
    if (!["ours", "theirs"].includes(side)) throw bad('side must be "ours" or "theirs"');
    const doc = await readDocx(store.readBlob("files", t.fileHash, ".docx"));
    const other = store.versions(matterId)
      .filter((v) => !v.renditionOf && (side === "theirs" ? v.side === "ours" : v.side === "theirs" && v.status === "active"));
    const { ranked, flags } = rankBases(doc, other);
    res.json({ bases: ranked.map(({ versionId, label, coverage }) => ({ versionId, label, coverage })), flags });
  }));

  app.get("/api/triage", (_req, res) => res.json(store.openTriage()));

  app.post("/api/triage/:id/assign", wrap(async (req, res) => {
    const t = store.triage(req.params.id);
    if (!t || t.status !== "open") return res.status(404).json({ error: "no open triage item" });
    const { matterId, side, baseId } = req.body || {};
    if (!store.matter(matterId)) return res.status(400).json({ error: "unknown matterId" });

    if (t.kind === "document") {
      if (!["ours", "theirs"].includes(side)) throw bad('side must be "ours" or "theirs"');
      const report = await registerDocument(store.readBlob("files", t.fileHash, ".docx"), { matterId, side, baseId, fileName: t.fileName }, ctx);
      store.closeTriage(t.id, { assignedTo: matterId, versionId: report.versionId, closedAt: new Date().toISOString() });
      audit({ event: "triage_assign", triageId: t.id, kind: "document", matterId, side, versionId: report.versionId, baseLabel: report.baseLabel ?? null });
      return res.json({ ...report, timeline: buildTimeline(store.versions(), report.versionId) });
    }

    const report = await processMessage(store.readBlob("mail", t.emlHash, ".eml"), { ...ctx, forceMatterId: matterId });
    store.closeTriage(t.id, { assignedTo: matterId, closedAt: new Date().toISOString() });
    audit({ event: "triage_assign", triageId: t.id, kind: "mail", matterId });
    res.json(report);
  }));

  return app;
}
