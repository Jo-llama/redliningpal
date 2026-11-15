import { test, before } from "node:test";
import assert from "node:assert/strict";
import { readFileSync, mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { createContext } from "../server/context.js";
import { processMessage } from "../server/intake.js";
import { readDocx } from "../server/docx.js";
import { identifyDocument, buildTimeline } from "../server/timeline.js";

const emails = new URL("./fixtures/emails/", import.meta.url);
const sample = async (p) => readDocx(readFileSync(new URL(`../samples/${p}`, import.meta.url)));
let ctx;

before(async () => {
  ctx = {
    ...createContext({
      configFile: new URL("../config/intake.example.json", import.meta.url).pathname,
      dataDir: mkdtempSync(path.join(tmpdir(), "redline-timeline-")),
    }),
    mock: true,
    notify: () => {},
  };
  for (const f of ["01_ours_v1.eml", "02_theirs_v2.eml", "03_theirs_v2-1_corrected.eml", "05_theirs_clean_copy_new_thread.eml"]) {
    await processMessage(readFileSync(new URL(f, emails)), ctx);
  }
});

const all = () => ctx.store.versions();
const byLabel = (l) => all().find((v) => v.matterId === "SPA-Nordlicht" && v.label === l);

// ---- identifying the open document ----------------------------------------

test("identify: the v2.1 file as received is an exact match", async () => {
  const r = identifyDocument(await sample("spa/spa_v2-1_theirs_corrected.docx"), all());
  assert.equal(r.status, "exact");
  assert.equal(r.versionId, byLabel("v2.1").id);
  assert.equal(r.matterId, "SPA-Nordlicht");
});

test("identify: v2.1 after the lawyer accepted and rejected some changes is still probably v2.1", () => {
  const v = byLabel("v2.1");
  const changed = v.afterTexts.map((t, i) => (t !== v.beforeTexts[i] ? i : -1)).filter((i) => i >= 0);
  assert.ok(changed.length >= 4, "fixture needs a few tracked paragraphs");
  const doc = { stamp: null, afterTexts: [...v.afterTexts], beforeTexts: [...v.beforeTexts] };
  doc.beforeTexts[changed[0]] = doc.afterTexts[changed[0]]; // accepted
  doc.beforeTexts[changed[1]] = doc.afterTexts[changed[1]]; // accepted
  doc.afterTexts[changed[2]] = doc.beforeTexts[changed[2]]; // rejected
  const r = identifyDocument(doc, all());
  assert.equal(r.status, "probable");
  assert.equal(r.versionId, v.id);
  assert.ok(r.score >= 0.6, String(r.score)); // every paragraph is still one v2.1 can produce
  assert.deepEqual(r.flags, []); // clear lead over v2, not flagged as ambiguous
});

test("identify: v2.1 with EVERY change accepted is v2.1, not v2's clean copy", () => {
  // regression: a symmetric (Jaccard) score preferred "v2 (clean)" here
  const v = byLabel("v2.1");
  const doc = { stamp: v.docStamp, afterTexts: [...v.afterTexts], beforeTexts: [...v.afterTexts] };
  const r = identifyDocument(doc, all());
  assert.equal(r.status, "probable");
  assert.equal(r.versionId, v.id);
  assert.equal(r.rendition, false);
});

test("identify: v2 with every change accepted is v2", () => {
  const v = byLabel("v2");
  const doc = { stamp: v.docStamp, afterTexts: [...v.afterTexts], beforeTexts: [...v.afterTexts] };
  // identical to the clean copy we filed as a rendition of v2
  assert.equal(identifyDocument(doc, all()).versionId, v.id);
});

test("identify: a clean copy resolves to the version it is a copy of", async () => {
  const r = identifyDocument(await sample("spa/spa_v2_theirs_clean.docx"), all());
  assert.equal(r.status, "exact");
  assert.equal(r.versionId, byLabel("v2").id);
  assert.equal(r.rendition, true);
});

test("identify: our own v1 is found and its stamp agrees", async () => {
  const doc = await sample("spa/spa_v1_ours.docx");
  const r = identifyDocument(doc, all());
  assert.equal(r.versionId, byLabel("v1").id);
  assert.deepEqual(r.flags, []);
});

test("identify: an unrelated contract is unknown and offers no confident candidate", async () => {
  const r = identifyDocument(await sample("nda_v2_counterparty_tracked.docx"), all());
  assert.equal(r.status, "unknown");
  assert.equal(r.versionId, null);
  assert.ok(r.candidates.every((c) => c.score > 0 && c.score < 0.6), "no zero-score noise, no confident match");
});

test("identify: text beats stamp - their v2.1 carries our v1 stamp but is not v1", async () => {
  const doc = await sample("spa/spa_v2-1_theirs_corrected.docx");
  assert.equal(doc.stamp, byLabel("v1").id, "fixture: counterparty kept our stamp");
  assert.equal(identifyDocument(doc, all()).versionId, byLabel("v2.1").id);
});

test("identify: empty registry -> unknown", () => {
  const r = identifyDocument({ stamp: null, afterTexts: ["a"], beforeTexts: ["a"] }, []);
  assert.equal(r.status, "unknown");
  assert.deepEqual(r.candidates, []);
});

// ---- timeline ---------------------------------------------------------------

test("timeline v2.1: base v1, supersedes v2, latest round, 1 untracked change", () => {
  const t = buildTimeline(all(), byLabel("v2.1").id);
  assert.equal(t.current.label, "v2.1");
  assert.equal(t.base.label, "v1");
  assert.deepEqual(t.supersedes.map((v) => v.label), ["v2"]);
  assert.deepEqual(t.supersededBy, []);
  assert.equal(t.newerExists, false);
  assert.deepEqual(t.untracked, { checked: true, count: 1, acknowledgedAt: null });
  assert.deepEqual(t.chain.map((c) => [c.label, c.role]), [["v1", "base"], ["v2", "superseded"], ["v2.1", "current"]]);
  assert.equal(t.chain.find((c) => c.label === "v2").renditions, 1);
});

test("timeline v2: superseded by v2.1, a newer version exists, no untracked changes", () => {
  const t = buildTimeline(all(), byLabel("v2").id);
  assert.deepEqual(t.supersededBy.map((v) => v.label), ["v2.1"]);
  assert.equal(t.newerExists, true);
  assert.equal(t.latest.label, "v2.1");
  assert.equal(t.untracked.count, 0);
});

test("timeline v1: no base, nothing superseded, untracked not checked", () => {
  const t = buildTimeline(all(), byLabel("v1").id);
  assert.equal(t.base, null);
  assert.deepEqual(t.supersedes, []);
  assert.equal(t.untracked.checked, false);
});

test("timeline of a rendition shows the version it is a copy of", () => {
  const clean = all().find((v) => v.renditionOf);
  const t = buildTimeline(all(), clean.id);
  assert.equal(t.current.label, "v2");
  assert.equal(t.openedRendition, true);
});

test("timeline: unknown id -> null", () => {
  assert.equal(buildTimeline(all(), "nope"), null);
});

test("timeline: loops and links into other matters are cut and flagged", () => {
  const mk = (id, matterId, extra = {}) => ({ id, label: id, matterId, side: "theirs", status: "active", receivedAt: id, ...extra });
  const versions = [
    mk("a", "M1", { supersedesId: "b", baseId: "x" }),
    mk("b", "M1", { supersedesId: "a" }),
    mk("x", "M2"),
  ];
  const t = buildTimeline(versions, "a");
  assert.deepEqual(t.supersedes.map((v) => v.id), ["b"]);
  assert.equal(t.base, null);
  assert.ok(t.flags.some((f) => /Loop/.test(f)));
  assert.ok(t.flags.some((f) => /not in this matter/.test(f)));
  assert.deepEqual(t.chain.map((c) => c.id), ["a", "b"]);
});

test("timeline and identify results carry no contract text", async () => {
  const doc = await sample("spa/spa_v2-1_theirs_corrected.docx");
  const payload = JSON.stringify([identifyDocument(doc, all()), ...all().map((v) => buildTimeline(all(), v.id))]);
  const texts = new Set(all().flatMap((v) => [...v.afterTexts, ...v.beforeTexts]).filter((t) => t.length > 20));
  assert.ok(texts.size > 10);
  for (const t of texts) assert.ok(!payload.includes(t), `leaked: ${t.slice(0, 40)}…`);
});
