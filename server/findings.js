// Playbook findings of a registered version, with how far each change has got in
// the document as it is now (Google Docs export or any later .docx of it).
// Pure: no I/O. Contract text goes to the lawyer's sidebar only, never to logs.

export const STATES = ["open", "countered", "accepted", "rejected", "changed"];

/**
 * Where is one reviewed change now?
 *  open       their change is still pending (tracked: their suggestion; untracked/compare: their text)
 *  countered  a suggestion on top of it that is neither theirs nor ours (our counter-proposal)
 *  accepted   their text, as plain text
 *  rejected   our text, as plain text (for their insertion: the text is gone)
 *  changed    none of the above (edited directly, or the paragraph is not found)
 *
 * @param unit  { source: "tracked"|"untracked"|"compare", before, after }  (normalised texts)
 * @param doc   { beforeTexts[], afterTexts[], paragraphs[{ hasChanges }] } from readDocx
 */
export function findingState(unit, doc) {
  const B = unit.before || "", A = unit.after || "";
  const paras = doc.afterTexts.map((after, i) => ({ before: doc.beforeTexts[i], after, pending: !!doc.paragraphs[i]?.hasChanges && doc.beforeTexts[i] !== after }));
  const plain = (t) => paras.some((p) => !p.pending && t && p.after === t);
  const anywhere = (t) => paras.some((p) => t && (p.after === t || p.before === t));
  const suggestion = (from, to) => paras.some((p) => p.pending && p.before === from && p.after === to);
  const onTopOf = (from) => paras.some((p) => p.pending && p.before === from && p.after !== A && p.after !== from);

  if (unit.source === "tracked") {
    if (suggestion(B, A)) return "open";
  } else if (A ? plain(A) : B && !anywhere(B)) {
    return "open";
  }
  // our suggestion on their text (accepted first) or on ours (theirs rejected first)
  if ((A && onTopOf(A)) || (B && onTopOf(B))) return "countered";
  // their deletion of a whole paragraph, answered with our suggestion to put it back
  if (!A && B && unit.source !== "tracked" && suggestion("", B)) return "countered";
  if (unit.source === "tracked") {
    if (A ? plain(A) : B && !anywhere(B)) return "accepted";
  }
  if (B ? plain(B) : A && !anywhere(A)) return "rejected";
  return "changed";
}

/**
 * The stored analysis of `version` with a state per change, for the sidebar.
 * @returns { versionId, playbook, mode, model, reviewedAt, harnessNotes, counts, progress, units[] }
 */
export function findingsFor(version, doc) {
  const a = version.analysis || {};
  const units = (a.units || []).map((u) => ({
    id: u.id, source: u.source, kind: u.kind, section: u.section || null,
    before: u.before || "", after: u.after || "", status: u.status || "new",
    state: findingState(u, doc),
    finding: {
      verdict: u.finding?.verdict, risk: u.finding?.risk, ruleIds: u.finding?.ruleIds || [],
      rationale: u.finding?.rationale || "", counterProposal: u.finding?.counterProposal ?? null,
      comment: u.finding?.comment || "", flags: u.finding?.flags || [],
    },
  }));
  const progress = Object.fromEntries(STATES.map((s) => [s, units.filter((u) => u.state === s).length]));
  return {
    versionId: version.id, playbook: a.playbook || null, mode: a.mode || null, model: a.model || null,
    reviewedAt: a.reviewedAt || null, harnessNotes: (a.harnessIssues || []).length,
    counts: a.counts || {}, progress, units,
  };
}
