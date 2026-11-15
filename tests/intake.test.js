import { test, before } from "node:test";
import assert from "node:assert/strict";
import { readFileSync, mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { createContext } from "../server/context.js";
import { processMessage } from "../server/intake.js";
import { readDocx, writeStamp } from "../server/docx.js";
import { classifyIntent } from "../server/intent.js";

const dir = new URL("./fixtures/emails/", import.meta.url);
const eml = (name) => readFileSync(new URL(name, dir));
let ctx;
const log = [];
const reports = {};

before(async () => {
  ctx = {
    ...createContext({
      configFile: new URL("../config/intake.example.json", import.meta.url).pathname,
      dataDir: mkdtempSync(path.join(tmpdir(), "redline-")),
    }),
    mock: true,
    notify: (t) => log.push(t),
  };
  // process the mailbox in order, like the poller would
  for (const f of [
    "01_ours_v1.eml", "02_theirs_v2.eml", "03_theirs_v2-1_corrected.eml", "04_duplicate_of_02.eml",
    "05_theirs_clean_copy_new_thread.eml", "06_no_attachment.eml", "07_unknown_sender.eml",
  ]) reports[f] = await processMessage(eml(f), ctx);
});

const versions = () => ctx.store.versions("SPA-Nordlicht");
const byLabel = (l) => versions().find((v) => v.label === l);

test("01 outbound draft registers our v1 with its stamp", () => {
  const r = reports["01_ours_v1.eml"];
  assert.equal(r.outcome, "processed");
  assert.equal(r.matterId, "SPA-Nordlicht");
  const v1 = byLabel("v1");
  assert.equal(v1.side, "ours");
  assert.equal(v1.id, "7f3a9c2e-5b1d-4c8a-9e6f-0a1b2c3d4e5f");
});

test("02 counterparty markup -> v2, base v1, routed by stamp + thread + tag", () => {
  const r = reports["02_theirs_v2.eml"];
  assert.equal(r.matterId, "SPA-Nordlicht");
  assert.ok(r.routing.auto);
  const ev = r.routing.candidates[0].evidence.join(" ");
  assert.match(ev, /stamp/);
  assert.match(ev, /thread/);
  const v2 = byLabel("v2");
  assert.equal(v2.side, "theirs");
  assert.equal(v2.baseId, byLabel("v1").id);
  assert.equal(v2.analysis.integrity.untracked, 0);
  assert.equal(v2.analysis.units.length, 12);
});

test("03 corrected version -> v2.1 supersedes v2, same base, untracked edit caught", () => {
  const r = reports["03_theirs_v2-1_corrected.eml"];
  assert.equal(r.intent, "corrected_version");
  const v2 = byLabel("v2"), v21 = byLabel("v2.1");
  assert.equal(v21.supersedesId, v2.id);
  assert.equal(v2.status, "superseded");
  assert.equal(v21.baseId, byLabel("v1").id);

  const a = v21.analysis;
  assert.equal(a.integrity.untracked, 1);
  const silent = a.units.find((u) => u.source === "untracked");
  assert.match(silent.before, /intent or fraud/);
  assert.equal(silent.after, "4.4 None of the limitations in this Clause 4 apply in case of intent.");
  assert.equal(silent.finding.risk, "high");
  assert.ok(["human_review", "reject"].includes(silent.finding.verdict));

  // delta: 8.1, 4.3 changed vs v2 + the silent edit; the rest carried over
  assert.equal(a.delta.newUnits, 3);
  assert.equal(a.delta.carriedOver, a.units.length - 3);
  assert.ok(a.units.filter((u) => u.status === "carried_over").every((u) => u.finding.reused));
});

test("04 duplicate delivery is ignored", () => {
  assert.equal(reports["04_duplicate_of_02.eml"].outcome, "duplicate");
});

test("05 clean copy in a new thread is routed by domain + text and filed as rendition", () => {
  const r = reports["05_theirs_clean_copy_new_thread.eml"];
  assert.equal(r.matterId, "SPA-Nordlicht");
  assert.equal(r.documents[0].outcome, "rendition");
  assert.equal(r.documents[0].versionLabel, "v2");
  assert.match(r.skipped[0].reason, /PDF/);
});

test("06 mail without document is linked to the matter via the thread", () => {
  const r = reports["06_no_attachment.eml"];
  assert.equal(r.outcome, "no_document");
  assert.equal(r.matterId, "SPA-Nordlicht");
});

test("07 unknown sender with unrelated contract goes to triage, assign re-processes it", async () => {
  const r = reports["07_unknown_sender.eml"];
  assert.equal(r.outcome, "triage");
  assert.equal(ctx.store.openTriage().length, 1);
  const again = await processMessage(eml("07_unknown_sender.eml"), { ...ctx, forceMatterId: "SPA-Polaris" });
  assert.equal(again.outcome, "processed");
  assert.equal(again.matterId, "SPA-Polaris");
});

test("notifications summarise what a lawyer needs to know", () => {
  assert.ok(log.some((l) => /v2\.1 .*replaces v2.*1 UNTRACKED.*3 new since v2/.test(l)), log.join("\n"));
});

test("stamp round-trip on a .docx", async () => {
  const buf = readFileSync(new URL("../samples/nda_v1_ours.docx", import.meta.url));
  assert.equal((await readDocx(buf)).stamp, null);
  const once = await writeStamp(buf, "abc-1");
  assert.equal((await readDocx(once)).stamp, "abc-1");
  const twice = await writeStamp(once, "abc-2");
  assert.equal((await readDocx(twice)).stamp, "abc-2");
});

test("intent rules: DE / EN / PL", () => {
  const c = (text) => classifyIntent({ subject: "", text }, true).intent;
  assert.equal(c("Anbei die korrigierte Fassung, bitte die vorherige ignorieren."), "corrected_version");
  assert.equal(c("W załączniku poprawiona wersja umowy."), "corrected_version");
  assert.equal(c("Attached the execution version for signature."), "final_for_signature");
  assert.equal(c("Our comments attached."), "new_version");
});

test("forwarded mails: counterparty's mail -> inbound, our own sent mail -> outbound", async () => {
  const mk = (orig, id) => Buffer.from([
    "From: Jo <buyer.counsel@halden-legal.example>", "To: redline-intake@halden-legal.example",
    "Subject: Fwd: [SPA-Nordlicht] something", `Message-ID: <${id}@x>`, "Date: Mon, 16 Nov 2026 10:00:00 +0000",
    "Content-Type: text/plain", "", "---------- Forwarded message ---------", `From: X <${orig}>`, "Subject: hi", "", "text",
  ].join("\r\n"));
  const a = await processMessage(mk("m.hughes@northwind-legal.example", "f1"), ctx);
  const b = await processMessage(mk("buyer.counsel@halden-legal.example", "f2"), ctx);
  assert.equal(a.direction, "inbound");
  assert.equal(a.from, "m.hughes@northwind-legal.example");
  assert.equal(b.direction, "outbound");
});
