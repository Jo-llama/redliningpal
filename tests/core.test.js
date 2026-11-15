import { test } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import JSZip from "jszip";
import { DOMParser } from "@xmldom/xmldom";

import { parseParagraphs, unitsFromTrackedChanges } from "../public/lib/ooxml.js";
import { diffParagraphs, attachSections, similarity } from "../public/lib/diff.js";
import { validateFindings, mockReview, guardrailHits } from "../public/lib/review.js";

const playbook = JSON.parse(readFileSync(new URL("../playbooks/mutual-nda.json", import.meta.url)));

async function docXml(name) {
  const buf = readFileSync(new URL(`../samples/${name}`, import.meta.url));
  const zip = await JSZip.loadAsync(buf);
  return zip.file("word/document.xml").async("string");
}

test("tracked changes: reconstructs before/after per paragraph", async () => {
  const paras = parseParagraphs(await docXml("nda_v2_counterparty_tracked.docx"), DOMParser);
  const units = unitsFromTrackedChanges(paras);
  const byStart = (s) => units.find((u) => (u.after || u.before).startsWith(s));

  assert.equal(units.length, 9);
  assert.match(byStart("1.1").before, /parties wish to evaluate/);
  assert.match(byStart("1.1").after, /parties intend to evaluate/);
  assert.match(byStart("5.1").after, /two \(2\) years after the last disclosure\.$/);
  assert.match(byStart("5.1").before, /trade secrets/);
  assert.equal(byStart("7.2").kind, "inserted");
  assert.equal(byStart("9. Non").kind, "inserted");
  assert.deepEqual(byStart("8.1").authors, ["Northwind Legal"]);
});

test("compare mode: clean v2 vs our v1 yields the same changes", async () => {
  const v1 = parseParagraphs(await docXml("nda_v1_ours.docx"), DOMParser).map((p) => p.after);
  const v2 = parseParagraphs(await docXml("nda_v2_counterparty_clean.docx"), DOMParser).map((p) => p.after);
  const units = diffParagraphs(v1, v2);

  assert.equal(units.length, 9);
  assert.equal(units.filter((u) => u.kind === "inserted").length, 3);
  assert.equal(units.filter((u) => u.kind === "modified").length, 6);
  for (const u of units) assert.equal(v2[u.paragraphIndex].replace(/\s+/g, " ").trim(), u.after);
});

test("compare mode: deleted paragraph is reported and anchored", () => {
  const units = diffParagraphs(["A one", "B two three", "C four"], ["A one", "C four"]);
  assert.equal(units.length, 1);
  assert.equal(units[0].kind, "deleted");
  assert.equal(units[0].before, "B two three");
  assert.equal(units[0].paragraphIndex, 1);
});

test("sections attach to nearest heading", async () => {
  const paras = parseParagraphs(await docXml("nda_v2_counterparty_tracked.docx"), DOMParser);
  const units = attachSections(unitsFromTrackedChanges(paras), paras.map((p) => p.after));
  assert.equal(units.find((u) => u.after.startsWith("5.1")).section, "5. Term");
});

test("guardrails catch red flags only in added text", () => {
  const residuals = { before: "x (d) independent.", after: "x (d) independent, or (e) retained in the unaided memory (residuals)." };
  assert.deepEqual(guardrailHits(playbook, residuals).map((h) => h.ruleId), ["NDA-04"]);
  // the counterparty *removing* a residuals clause must not trigger
  assert.equal(guardrailHits(playbook, { before: residuals.after, after: residuals.before }).length, 0);
});

test("validator: fills gaps, drops junk, guardrails only escalate", () => {
  const units = [
    { id: "U1", before: "a", after: "b" },
    { id: "U2", before: "", after: "Neither party shall solicit or employ any employee." },
    { id: "U3", before: "x", after: "y" },
  ];
  const raw = {
    findings: [
      { unit_id: "U1", verdict: "accept", risk: "low", rule_ids: ["NDA-99"], rationale: "ok", counter_proposal: null, comment: "" },
      { unit_id: "U2", verdict: "accept", risk: "low", rule_ids: [], rationale: "fine", counter_proposal: null, comment: "" },
      { unit_id: "U2", verdict: "reject", risk: "high", rule_ids: [], rationale: "dupe", counter_proposal: null, comment: "" },
      { unit_id: "U9", verdict: "reject", risk: "high", rule_ids: [], rationale: "ghost", counter_proposal: null, comment: "" },
    ],
  };
  const { findings, issues } = validateFindings(raw, units, playbook);
  assert.equal(findings.length, 3);
  assert.deepEqual(findings[0].ruleIds, []);           // unknown rule removed
  assert.equal(findings[1].verdict, "reject");          // escalated by NDA-05 red flag
  assert.ok(findings[1].ruleIds.includes("NDA-05"));
  assert.equal(findings[2].verdict, "human_review");    // missing -> human review
  assert.ok(issues.some((i) => i.includes("unknown unit U9")));
  assert.ok(issues.some((i) => i.includes("duplicate")));
});

test("mock reviewer produces a valid finding for every sample change", async () => {
  const paras = parseParagraphs(await docXml("nda_v2_counterparty_tracked.docx"), DOMParser);
  const units = attachSections(unitsFromTrackedChanges(paras), paras.map((p) => p.after));
  const { findings, issues } = validateFindings(mockReview(playbook, units), units, playbook);
  assert.equal(findings.length, units.length);
  assert.equal(issues.filter((i) => i.includes("no finding")).length, 0);
  const verdictFor = (s) => findings[units.findIndex((u) => u.after.startsWith(s))].verdict;
  assert.equal(verdictFor("3.1"), "reject");  // residuals
  assert.equal(verdictFor("9.1"), "reject");  // non-solicit
});

test("similarity sanity", () => {
  assert.ok(similarity("The term is five years", "The term is two years") > 0.6);
  assert.ok(similarity("Governing law", "Non-solicitation of staff") < 0.2);
});

import { wordDiff, minimalEdit, suggestionEdit } from "../public/lib/diff.js";

test("minimalEdit keeps shared prefix/suffix", () => {
  const e = minimalEdit("continue for five (5) years after", "continue for three (3) years after");
  assert.equal(e.oldMiddle, "five (5)");
  assert.equal(e.newMiddle, "three (3)");
});

test("wordDiff marks insertions and deletions", () => {
  const d = wordDiff("governed by German law.", "governed by the laws of England.");
  assert.ok(d.some((s) => s.op === "-" && s.text.includes("German")));
  assert.ok(d.some((s) => s.op === "+" && s.text.includes("England")));
});

test("wordDiff groups a changed phrase: all deleted words, then all inserted words", () => {
  const show = (d) => d.map((p) => (p.op === "=" ? p.text : `[${p.op}${p.text}]`)).join("");
  const b = "in which case the Seller is liable from the first euro.";
  const a = "in which case the Seller is liable only for the excess.";
  const d = wordDiff(b, a);
  assert.equal(d.map((p) => (p.op === "+" ? "" : p.text)).join(""), b, "deleted + shared = before");
  assert.equal(d.map((p) => (p.op === "-" ? "" : p.text)).join(""), a, "inserted + shared = after");
  for (let k = 1; k < d.length; k++) assert.ok(!(d[k - 1].op === "+" && d[k].op === "-"), `no insertion before a deletion: ${show(d)}`);
  assert.equal(show(wordDiff("For three (3) years", "For two (2) years")), "For [-three (3)][+two (2)] years");
  assert.equal(show(wordDiff("a b c d", "a d")), "a [-b c ]d", "pure deletion unchanged");
});

// ---- suggestionEdit: the smallest suggestion a lawyer applies by hand (Google Docs)

test("suggestionEdit: a deleted 'or fraud' becomes an insertion between 'intent' and '.'", () => {
  const e = suggestionEdit("4.4 None of the limitations in this Clause 4 apply in case of intent.",
    "4.4 None of the limitations in this Clause 4 apply in case of intent or fraud.");
  assert.equal(e.kind, "insert");
  assert.equal(e.newText, " or fraud");
  assert.equal(e.hint, "Insert “or fraud” after “…in case of intent”");
  assert.equal(e.after, ".");
  assert.ok(e.before.endsWith("in case of intent"));
});

test("suggestionEdit: deletion reads as whole words, at the place a markup would show it", () => {
  const e = suggestionEdit("accurate on the Signing Date and on the Closing Date.", "accurate on the Signing Date.");
  assert.equal(e.kind, "delete");
  assert.equal(e.oldText, " and on the Closing Date");
  assert.equal(e.hint, "Delete “and on the Closing Date”");
  assert.equal(e.before, "accurate on the Signing Date");
  assert.equal(e.after, ".");
});

test("suggestionEdit: replacements stay whole words, not fragments like 'five (5'", () => {
  const e = suggestionEdit("continue for five (5) years after", "continue for three (3) years after");
  assert.deepEqual([e.kind, e.oldText, e.newText], ["replace", "five (5)", "three (3)"]);
  assert.equal(e.hint, "Replace “five (5)” with “three (3)”");
  const n = suggestionEdit("claims below EUR 100,000 (de minimis).", "claims below EUR 25,000 (de minimis).");
  assert.deepEqual([n.oldText, n.newText, n.before, n.after], ["100,000", "25,000", "claims below EUR ", " (de minimis)."]);
});

test("suggestionEdit: whitespace differences are no change; start of paragraph; too long -> whole", () => {
  assert.equal(suggestionEdit("same  text ", "same text").kind, "none");
  const start = suggestionEdit("Seller warrants.", "The Seller warrants.");
  assert.deepEqual([start.kind, start.newText, start.hint], ["insert", "The ", "Insert “The” at the start of the paragraph"]);
  assert.equal(suggestionEdit("Price.", `Price; ${"x ".repeat(120)}`).kind, "whole");
  assert.equal(suggestionEdit("", "4.4 restored paragraph").kind, "whole");
});

test("suggestionEdit: context is a few words, so the spot can be found again", () => {
  const long = "1 2 3 4 5 6 7 8 9 10 claims below EUR 100,000 are excluded from 11 12 13 14 15 16 17 18";
  const e = suggestionEdit(long, long.replace("100,000", "25,000"));
  assert.equal(e.before, "5 6 7 8 9 10 claims below EUR ".slice(-e.before.length));
  assert.ok(e.before.split(" ").filter(Boolean).length <= 6);
  assert.ok(e.after.split(" ").filter(Boolean).length <= 6);
});
