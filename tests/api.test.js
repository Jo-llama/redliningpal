import { test, before, after } from "node:test";
import assert from "node:assert/strict";
import { readFileSync, mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import http from "node:http";
import path from "node:path";
import JSZip from "jszip";
import { createContext } from "../server/context.js";
import { processMessage } from "../server/intake.js";
import { readDocx } from "../server/docx.js";
import { createApp } from "../server/app.js";

const root = new URL("..", import.meta.url).pathname;
const file = (p) => readFileSync(new URL(`../samples/${p}`, import.meta.url));
let ctx, server, base;
const auditLog = [];

before(async () => {
  ctx = {
    ...createContext({
      configFile: new URL("../config/intake.example.json", import.meta.url).pathname,
      dataDir: mkdtempSync(path.join(tmpdir(), "redline-api-")),
    }),
    mock: true,
    notify: () => {},
  };
  for (const f of ["01_ours_v1.eml", "02_theirs_v2.eml", "03_theirs_v2-1_corrected.eml", "05_theirs_clean_copy_new_thread.eml"]) {
    await processMessage(readFileSync(new URL(`./fixtures/emails/${f}`, import.meta.url)), ctx);
  }
  const app = createApp(ctx, { root, preview: false, audit: (e) => auditLog.push(e) });
  server = http.createServer(app).listen(0);
  await new Promise((r) => server.once("listening", r));
  base = `http://127.0.0.1:${server.address().port}`;
});

after(() => server.close());

const api = async (method, url, body, headers = {}) => {
  const isBuf = Buffer.isBuffer(body);
  const res = await fetch(base + url, {
    method,
    headers: body === undefined ? headers : { "content-type": isBuf ? "application/octet-stream" : "application/json", ...headers },
    body: body === undefined ? undefined : isBuf ? body : JSON.stringify(body),
  });
  return { status: res.status, body: await res.json().catch(() => null) };
};
const byLabel = (l) => ctx.store.versions("SPA-Nordlicht").find((v) => v.label === l);
const textsOf = async (buf) => {
  const d = await readDocx(buf);
  return { afterTexts: d.afterTexts, beforeTexts: d.beforeTexts, stamp: d.stamp };
};

/** v2.1 with one paragraph added: our counter-draft, not a registered version. */
async function counterDraft() {
  const zip = await JSZip.loadAsync(file("spa/spa_v2-1_theirs_corrected.docx"));
  const xml = await zip.file("word/document.xml").async("string");
  const para = "<w:p><w:r><w:t>15. Escrow: 10 percent of the purchase price is held for 18 months.</w:t></w:r></w:p>";
  zip.file("word/document.xml", xml.replace(/<w:sectPr[\s>]/, (m) => para + m));
  return zip.generateAsync({ type: "nodebuffer" });
}

// ---- what is open in Word ------------------------------------------------------

test("identify: v2.1 open in Word -> exact, with timeline", async () => {
  const r = await api("POST", "/api/documents/identify", await textsOf(file("spa/spa_v2-1_theirs_corrected.docx")));
  assert.equal(r.status, 200);
  assert.equal(r.body.status, "exact");
  assert.equal(r.body.timeline.current.label, "v2.1");
  assert.equal(r.body.timeline.base.label, "v1");
  assert.deepEqual(r.body.timeline.supersedes.map((v) => v.label), ["v2"]);
  const { afterTexts } = await textsOf(file("spa/spa_v2-1_theirs_corrected.docx"));
  assert.deepEqual(auditLog.at(-1), { event: "identify", status: "exact", versionId: byLabel("v2.1").id, score: 1, paragraphs: afterTexts.length });
});

test("identify: unknown document -> no timeline, candidates only", async () => {
  const r = await api("POST", "/api/documents/identify", await textsOf(file("nda_v2_counterparty_tracked.docx")));
  assert.equal(r.body.status, "unknown");
  assert.equal(r.body.timeline, null);
});

test("identify: malformed input -> 400", async () => {
  assert.equal((await api("POST", "/api/documents/identify", {})).status, 400);
  assert.equal((await api("POST", "/api/documents/identify", { afterTexts: [1], beforeTexts: [1] })).status, 400);
  assert.equal((await api("POST", "/api/documents/identify", { afterTexts: ["a", "b"], beforeTexts: ["a"] })).status, 400);
});

test("timeline route: known version -> 200, unknown -> 404", async () => {
  assert.equal((await api("GET", `/api/versions/${byLabel("v2").id}/timeline`)).body.newerExists, true);
  assert.equal((await api("GET", "/api/versions/nope/timeline")).status, 404);
});

test("base file can be downloaded for 'Open base version'", async () => {
  const res = await fetch(`${base}/api/versions/${byLabel("v1").id}/file`);
  assert.equal(res.status, 200);
  const doc = await readDocx(Buffer.from(await res.arrayBuffer()));
  assert.deepEqual(doc.afterTexts, byLabel("v1").afterTexts);
});

// ---- untracked changes ------------------------------------------------------------

test("untracked: v2.1 has one card, v2 none", async () => {
  const r = await api("GET", `/api/versions/${byLabel("v2.1").id}/untracked`);
  assert.equal(r.body.units.length, 1);
  assert.match(r.body.units[0].before, /intent or fraud/);
  assert.ok(Number.isInteger(r.body.units[0].paragraphIndex), "the pane needs it to locate the card");
  assert.ok(["human_review", "reject"].includes(r.body.units[0].finding.verdict));
  assert.equal((await api("GET", `/api/versions/${byLabel("v2").id}/untracked`)).body.units.length, 0);
});

test("ack-untracked: refused without untracked changes, recorded and audited otherwise", async () => {
  assert.equal((await api("POST", `/api/versions/${byLabel("v2").id}/ack-untracked`)).status, 409);
  const v21 = byLabel("v2.1");
  const r = await api("POST", `/api/versions/${v21.id}/ack-untracked`);
  assert.equal(r.status, 200);
  const t = (await api("GET", `/api/versions/${v21.id}/timeline`)).body;
  assert.equal(t.untracked.acknowledgedAt, r.body.acknowledgedAt);
  assert.deepEqual(auditLog.at(-1), { event: "ack_untracked", versionId: v21.id, matterId: "SPA-Nordlicht", untracked: 1 });
});

// ---- unknown document: triage, a human decides ----------------------------------------

test("triage: a registered document is refused (409), nothing filed", async () => {
  const before = ctx.store.openTriage().length;
  const r = await api("POST", "/api/documents/triage", file("spa/spa_v2-1_theirs_corrected.docx"));
  assert.equal(r.status, 409);
  assert.equal(r.body.versionId, byLabel("v2.1").id);
  assert.equal(ctx.store.openTriage().length, before);
});

test("triage: not a .docx -> 400", async () => {
  assert.equal((await api("POST", "/api/documents/triage", Buffer.from("hello"))).status, 400);
});

test("triage: unknown NDA is filed for a decision, not registered; assign needs side, then registers", async () => {
  const versionsBefore = ctx.store.versions().length;
  const r = await api("POST", "/api/documents/triage", file("nda_v2_counterparty_tracked.docx"), { "x-file-name": "NDA <x>.docx" });
  assert.equal(r.status, 201);
  const id = r.body.triageId;
  assert.ok(r.body.matters.some((m) => m.id === "SPA-Polaris"));
  assert.equal(ctx.store.versions().length, versionsBefore, "nothing registered before a human decides");
  const t = ctx.store.triage(id);
  assert.equal(t.kind, "document");
  assert.equal(t.fileName, "NDA _x_.docx");

  assert.equal((await api("POST", `/api/triage/${id}/assign`, { matterId: "SPA-Polaris" })).status, 400);
  assert.equal((await api("POST", `/api/triage/${id}/assign`, { matterId: "nope", side: "theirs" })).status, 400);

  const a = await api("POST", `/api/triage/${id}/assign`, { matterId: "SPA-Polaris", side: "theirs" });
  assert.equal(a.status, 200, JSON.stringify(a.body));
  assert.equal(a.body.outcome, "registered");
  assert.equal(a.body.timeline.current.label, "v1");
  assert.ok(a.body.flags.some((f) => /No version of ours/.test(f)));
  assert.equal(ctx.store.triage(id).status, "closed");
  assert.equal((await api("POST", `/api/triage/${id}/assign`, { matterId: "SPA-Polaris", side: "theirs" })).status, 404);
  assert.deepEqual(auditLog.at(-1), {
    event: "triage_assign", triageId: id, kind: "document", matterId: "SPA-Polaris", side: "theirs", versionId: a.body.versionId, baseLabel: null,
  });
});

test("triage: our counter-draft on v2.1 - suggested base, wrong-side base refused, confirmed base used", async () => {
  const buf = await counterDraft();
  const idn = await api("POST", "/api/documents/identify", await textsOf(buf));
  assert.equal(idn.body.status, "probable", "close to v2.1 but not registered");

  const r = await api("POST", "/api/documents/triage", buf);
  assert.equal(r.status, 201);
  const id = r.body.triageId;

  const b = await api("GET", `/api/triage/${id}/bases?matterId=SPA-Nordlicht&side=ours`);
  assert.equal(b.body.bases[0].label, "v2.1");
  assert.equal((await api("GET", `/api/triage/${id}/bases?matterId=SPA-Nordlicht&side=sideways`)).status, 400);

  const wrong = await api("POST", `/api/triage/${id}/assign`, { matterId: "SPA-Nordlicht", side: "ours", baseId: byLabel("v1").id });
  assert.equal(wrong.status, 400);
  assert.equal(ctx.store.triage(id).status, "open", "a refused assignment leaves the item open");

  const a = await api("POST", `/api/triage/${id}/assign`, { matterId: "SPA-Nordlicht", side: "ours", baseId: byLabel("v2.1").id });
  assert.equal(a.status, 200, JSON.stringify(a.body));
  assert.equal(a.body.versionLabel, "v3");
  assert.equal(a.body.timeline.base.label, "v2.1");
  assert.equal(a.body.timeline.latest.label, "v3");
  assert.equal((await api("GET", `/api/versions/${byLabel("v2.1").id}/timeline`)).body.newerExists, true);

  // the same draft again is now a registered document
  assert.equal((await api("POST", "/api/documents/triage", buf)).status, 409);
});

// ---- privacy --------------------------------------------------------------------------

test("audit lines carry ids and counts only, never contract text", async () => {
  assert.ok(auditLog.length >= 6);
  const lines = JSON.stringify(auditLog);
  const texts = ctx.store.versions().flatMap((v) => [...v.afterTexts, ...v.beforeTexts]).filter((t) => t.length > 20);
  for (const t of new Set(texts)) assert.ok(!lines.includes(t), `leaked: ${t.slice(0, 40)}…`);
});

test("timeline payloads carry no contract text", async () => {
  const payload = JSON.stringify(await Promise.all(ctx.store.versions().map((v) => api("GET", `/api/versions/${v.id}/timeline`))));
  const texts = ctx.store.versions().flatMap((v) => [...v.afterTexts, ...v.beforeTexts]).filter((t) => t.length > 20);
  for (const t of new Set(texts)) assert.ok(!payload.includes(t), `leaked: ${t.slice(0, 40)}…`);
});
