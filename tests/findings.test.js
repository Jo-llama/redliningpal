import { test } from "node:test";
import assert from "node:assert/strict";
import { findingState, findingsFor } from "../server/findings.js";

/** paragraphs as [before, after] (a suggestion pending where they differ) */
const doc = (...paras) => ({
  beforeTexts: paras.map(([b]) => b),
  afterTexts: paras.map(([, a]) => a),
  paragraphs: paras.map(([b, a]) => ({ hasChanges: b !== a })),
});
const OURS = "4.1 The Seller is not liable for claims below EUR 25,000.";
const THEIRS = "4.1 The Seller is not liable for claims below EUR 100,000.";
const COUNTER = "4.1 The Seller is not liable for claims below EUR 50,000.";
const other = ["1. Sale and Purchase", "1. Sale and Purchase"];
const tracked = { source: "tracked", before: OURS, after: THEIRS };

test("tracked change: open, accepted, rejected, changed", () => {
  assert.equal(findingState(tracked, doc(other, [OURS, THEIRS])), "open");
  assert.equal(findingState(tracked, doc(other, [THEIRS, THEIRS])), "accepted");
  assert.equal(findingState(tracked, doc(other, [OURS, OURS])), "rejected");
  assert.equal(findingState(tracked, doc(other, ["4.1 rewritten by hand", "4.1 rewritten by hand"])), "changed");
  assert.equal(findingState(tracked, doc(other)), "changed", "paragraph not found");
});

test("tracked change: our counter-proposal after rejecting theirs, or on top of theirs", () => {
  assert.equal(findingState(tracked, doc(other, [OURS, COUNTER])), "countered", "rejected theirs, suggested ours");
  assert.equal(findingState(tracked, doc(other, [THEIRS, COUNTER])), "countered", "accepted theirs, suggested ours");
});

test("tracked insertion and deletion of whole paragraphs", () => {
  const ins = { source: "tracked", before: "", after: "9.1 New employee clause." };
  assert.equal(findingState(ins, doc(other, ["", ins.after])), "open");
  assert.equal(findingState(ins, doc(other, [ins.after, ins.after])), "accepted");
  assert.equal(findingState(ins, doc(other)), "rejected");
  const del = { source: "tracked", before: "7. MAC clause.", after: "" };
  assert.equal(findingState(del, doc(other, [del.before, ""])), "open");
  assert.equal(findingState(del, doc(other)), "accepted");
  assert.equal(findingState(del, doc(other, [del.before, del.before])), "rejected");
  assert.equal(findingState(del, doc(other, [del.before, "7. MAC clause, narrowed."])), "countered");
});

test("untracked change: open while their text stands, countered by our suggestion, rejected once ours is back", () => {
  const silent = { source: "untracked", before: "4.4 … intent or fraud.", after: "4.4 … intent." };
  assert.equal(findingState(silent, doc(other, [silent.after, silent.after])), "open");
  assert.equal(findingState(silent, doc(other, [silent.after, silent.before])), "countered", "the minimal edit, pending");
  assert.equal(findingState(silent, doc(other, [silent.before, silent.before])), "rejected");
  const gone = { source: "untracked", before: "11. Confidentiality.", after: "" };
  assert.equal(findingState(gone, doc(other)), "open");
  assert.equal(findingState(gone, doc(other, ["", gone.before])), "countered", "suggested back");
  assert.equal(findingState(gone, doc(other, [gone.before, gone.before])), "rejected");
});

test("findingsFor: whitelisted fields, progress per state, harness notes as a count", () => {
  const version = {
    id: "v-2", analysis: {
      playbook: "spa-share-deal@2026-11", mode: "mock", model: null, reviewedAt: "2026-11-12T10:00:00Z",
      harnessIssues: ["U1: guardrail escalated negotiate -> reject (fraud)"], counts: { reject: 1, negotiate: 1 },
      units: [
        { id: "U1", source: "tracked", kind: "modified", section: "4. Limitation", before: OURS, after: THEIRS, cacheKey: "x", status: "new",
          finding: { verdict: "reject", risk: "high", ruleIds: ["SPA-04"], rationale: "r", counterProposal: OURS, comment: "c", flags: [], unitId: "U1", reused: true } },
        { id: "U2", source: "tracked", kind: "inserted", section: null, before: "", after: "9.1 New.", cacheKey: "y", status: "carried_over",
          finding: { verdict: "negotiate", risk: "medium", ruleIds: [], rationale: "", counterProposal: null, comment: "", flags: [] } },
      ],
    },
  };
  const r = findingsFor(version, doc(other, [OURS, THEIRS]));
  assert.deepEqual(r.progress, { open: 1, countered: 0, accepted: 0, rejected: 1, changed: 0 });
  assert.equal(r.harnessNotes, 1);
  assert.deepEqual(Object.keys(r.units[0]).sort(), ["after", "before", "finding", "id", "kind", "section", "source", "state", "status"]);
  assert.deepEqual(Object.keys(r.units[0].finding).sort(), ["comment", "counterProposal", "flags", "rationale", "risk", "ruleIds", "verdict"]);
  assert.equal(r.units[1].status, "carried_over");
  assert.deepEqual(findingsFor({ id: "v-1" }, doc(other)).units, [], "a version without analysis");
});
