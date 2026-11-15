import { test } from "node:test";
import assert from "node:assert/strict";
import { versionView, gateClosed, untrackedUnits } from "../public/lib/version-view.js";

const node = (label, side, extra = {}) => ({
  id: `id-${label}`, label, side, status: "active", receivedAt: "2026-11-13T09:00:00Z", hasTrackedChanges: true, ...extra,
});
const timeline = (extra = {}) => ({
  matterId: "SPA-Nordlicht",
  current: node("v2.1", "theirs"),
  openedRendition: false,
  base: node("v1", "ours", { receivedAt: "2026-11-10T09:00:00Z" }),
  supersedes: [node("v2", "theirs", { status: "superseded" })],
  supersededBy: [],
  newerExists: false,
  latest: node("v2.1", "theirs"),
  untracked: { checked: true, count: 1, acknowledgedAt: null },
  chain: [
    { ...node("v1", "ours"), role: "base", renditions: 0 },
    { ...node("v2", "theirs"), role: "superseded", renditions: 1 },
    { ...node("v2.1", "theirs"), role: "current", renditions: 0 },
  ],
  flags: [],
  ...extra,
});

test("view: exact v2.1 shows base, supersedes, chain and the untracked badge", () => {
  const v = versionView({ status: "exact", versionId: "id-v2.1", score: 1, flags: [], timeline: timeline() });
  assert.equal(v.known, true);
  assert.equal(v.title, "v2.1 · counterparty");
  assert.equal(v.match, "matches the registered file");
  assert.equal(v.base.id, "id-v1");
  assert.equal(v.base.text, "v1 (ours, 2026-11-10)");
  assert.equal(v.supersedes, "v2");
  assert.equal(v.warning, null);
  assert.equal(v.untracked.text, "1 change without track changes");
  assert.equal(v.gateClosed, true);
  assert.equal(v.source, "tracked");
  assert.match(v.chain[1].text, /\+1 clean copy$/);
  assert.equal(v.chain.find((c) => c.current).id, "id-v2.1");
});

test("view: probable match says so, with the score", () => {
  const v = versionView({ status: "probable", versionId: "id-v2.1", score: 0.93, flags: ["Close match with v2"], timeline: timeline() });
  assert.equal(v.probable, true);
  assert.match(v.match, /93%/);
  assert.deepEqual(v.flags, ["Close match with v2"]);
});

test("view: superseded version warns that it is outdated", () => {
  const t = timeline({ current: node("v2", "theirs", { status: "superseded" }), supersedes: [], supersededBy: [node("v2.1", "theirs")], newerExists: true });
  assert.match(versionView({ status: "exact", score: 1, timeline: t }).warning, /^Superseded by v2\.1 .*outdated/);
});

test("view: a newer round (not a correction) is pointed out", () => {
  const t = timeline({ newerExists: true, latest: node("v3", "ours") });
  assert.match(versionView({ status: "exact", score: 1, timeline: t }).warning, /^A newer version exists: v3 \(ours/);
});

test("view: clean copy and a clean document start in compare mode", () => {
  const t = timeline({ current: node("v2", "theirs", { hasTrackedChanges: false }), openedRendition: true, untracked: { checked: false, count: 0, acknowledgedAt: null } });
  const v = versionView({ status: "exact", score: 1, rendition: true, timeline: t });
  assert.equal(v.match, "clean copy of v2");
  assert.equal(v.source, "compare");
  assert.equal(v.untracked.text, null);
  assert.equal(v.gateClosed, false);
});

test("view: review mode follows the open document, not the registered version", () => {
  // regression: a clean copy of v2 resolves to v2, which has tracked changes -> tracked mode found nothing
  const t = timeline({ current: node("v2", "theirs", { hasTrackedChanges: true }), openedRendition: true });
  assert.equal(versionView({ status: "exact", score: 1, rendition: true, timeline: t }, { hasTrackedChanges: false }).source, "compare");
  // a markup where every change was already accepted is compared with the base as well
  assert.equal(versionView({ status: "probable", score: 1, timeline: timeline() }, { hasTrackedChanges: false }).source, "compare");
  assert.equal(versionView({ status: "exact", score: 1, timeline: timeline() }, { hasTrackedChanges: true }).source, "tracked");
});

test("view: unknown document lists candidates as percentages", () => {
  const v = versionView({ status: "unknown", candidates: [{ matterId: "SPA-Polaris", label: "v1", score: 0.12 }], flags: [] });
  assert.equal(v.known, false);
  assert.deepEqual(v.candidates, [{ matterId: "SPA-Polaris", text: "SPA-Polaris: v1 (12% of paragraphs match)" }]);
});

test("gate: closed only while untracked changes are unacknowledged", () => {
  assert.equal(gateClosed(timeline()), true);
  assert.equal(gateClosed(timeline({ untracked: { checked: true, count: 1, acknowledgedAt: "2026-11-15T10:00:00Z" } })), false);
  assert.equal(gateClosed(timeline({ untracked: { checked: true, count: 0, acknowledgedAt: null } })), false);
  assert.equal(gateClosed(null), false);
});

test("untracked units: located by current text nearest the stored index, behave like compare units", () => {
  const parsed = ["1. Intro", "4.4 Short text.", "x", "4.4 Short text."].map((after, index) => ({ index, after, before: after }));
  const server = [
    { id: "U7", kind: "modified", section: "4.4", paragraphIndex: 3, before: "4.4 Long text.", after: "4.4 Short text.", finding: { verdict: "human_review" } },
    { id: "U8", kind: "deleted", section: "5", paragraphIndex: 2, before: "5. Gone", after: "", finding: { verdict: "human_review" } },
    { id: "U9", kind: "modified", section: "6", before: "6. a", after: "6. edited by hand since", finding: { verdict: "reject" } },
  ];
  const [a, b, c] = untrackedUnits(server, parsed);
  assert.equal(a.paragraphIndex, 3, "nearest of two identical paragraphs");
  assert.equal(a.source, "compare");
  assert.equal(a.untracked, true);
  assert.deepEqual(a.expect, { before: "4.4 Short text.", after: "4.4 Short text." });
  assert.equal(b.paragraphIndex, 2, "deleted paragraph: stored index");
  assert.equal(c.paragraphIndex, 0, "no text match and no index: start of document");
  assert.equal(a.id, "X-U7");
});
