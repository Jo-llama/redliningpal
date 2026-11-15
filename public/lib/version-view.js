// What the task pane shows about the open document's version, and what it
// allows. Pure: runs in the browser and in Node tests. No contract text in,
// none out, except in untrackedUnits (review cards, never logged).

import { norm } from "./ooxml.js";
import { suggestionEdit, wordDiff } from "./diff.js";

const date = (iso) => (iso ? String(iso).slice(0, 10) : "");
const ref = (n) => (n ? `${n.label} (${n.side}, ${date(n.receivedAt)})` : null);

/** "Accept all" stays blocked while untracked changes exist and nobody has acknowledged them. */
export function gateClosed(timeline) {
  const u = timeline?.untracked;
  return !!u && u.count > 0 && !u.acknowledgedAt;
}

/**
 * @param identify  response of POST /api/documents/identify (or a triage/assign result
 *                  shaped like it: { status, versionId, score, rendition, flags, timeline })
 * @param opened    { hasTrackedChanges } of the document open in Word: the review mode
 *                  follows what is open (a clean copy of a markup has none), not the registry
 */
export function versionView(identify, opened = {}) {
  const t = identify?.timeline;
  if (!identify || identify.status === "unknown" || !t) {
    return {
      known: false,
      candidates: (identify?.candidates || []).map((c) => ({
        matterId: c.matterId, text: `${c.matterId}: ${c.label} (${Math.round(c.score * 100)}% of paragraphs match)`,
      })),
      flags: identify?.flags || [],
    };
  }

  const cur = t.current;
  let match = identify.status === "exact" ? "matches the registered file" : `probable match (${Math.round(identify.score * 100)}% of paragraphs)`;
  if (identify.rendition || t.openedRendition) match = `clean copy of ${cur.label}`;

  let warning = null;
  if (t.supersededBy.length) {
    const last = t.supersededBy[t.supersededBy.length - 1];
    warning = `Superseded by ${ref(last)}. You are looking at an outdated version.`;
  } else if (t.newerExists) {
    warning = `A newer version exists: ${ref(t.latest)}.`;
  }

  return {
    known: true,
    versionId: cur.id,
    matterId: t.matterId,
    title: `${cur.label} · ${cur.side === "ours" ? "our version" : "counterparty"}`,
    match,
    probable: identify.status === "probable",
    warning,
    base: t.base ? { id: t.base.id, text: ref(t.base) } : null,
    supersedes: t.supersedes.length ? t.supersedes.map((s) => s.label).join(" ← ") : null,
    chain: t.chain.map((c) => ({
      id: c.id,
      text: `${c.label} · ${c.side} · ${date(c.receivedAt)}${c.renditions ? ` · +${c.renditions} clean cop${c.renditions > 1 ? "ies" : "y"}` : ""}`,
      role: c.role, status: c.status, current: c.role === "current",
    })),
    untracked: {
      count: t.untracked.count,
      acknowledgedAt: t.untracked.acknowledgedAt,
      text: t.untracked.count ? `${t.untracked.count} change${t.untracked.count > 1 ? "s" : ""} without track changes` : null,
    },
    gateClosed: gateClosed(t),
    // review mode the pane should start in
    source: (opened.hasTrackedChanges ?? cur.hasTrackedChanges) ? "tracked" : "compare",
    flags: [...(identify.flags || []), ...(t.flags || [])],
  };
}

/**
 * Server-side untracked units -> pane review units. They are plain text in the
 * document (no <w:ins>/<w:del>), so they behave like compare-mode units:
 * "Reject" writes our base text back as a tracked change.
 * Locate by current text nearest to the stored paragraph index.
 */
export function untrackedUnits(serverUnits, parsedParas) {
  return serverUnits.map((u) => {
    const hint = Number.isInteger(u.paragraphIndex) ? u.paragraphIndex : 0;
    const hits = u.after ? parsedParas.filter((p) => norm(p.after) === u.after).map((p) => p.index) : [];
    const paragraphIndex = hits.length ? hits.reduce((b, i) => (Math.abs(i - hint) < Math.abs(b - hint) ? i : b)) : Math.max(0, hint);
    return {
      id: `X-${u.id}`, source: "compare", untracked: true, kind: u.kind, section: u.section,
      before: u.before, after: u.after, paragraphIndex,
      expect: { before: u.after, after: u.after }, done: false,
      finding: u.finding,
    };
  });
}

// ---- playbook review (Google Docs sidebar) ------------------------------------------------

const VERDICT_ORDER = { reject: 0, human_review: 1, negotiate: 2, accept: 3 };
const STATE_TEXT = {
  open: "Open",
  countered: "Our counter-proposal is in",
  accepted: "Accepted",
  rejected: "Rejected",
  changed: "Changed since the review: check by hand",
};
const HANDLED = new Set(["countered", "accepted", "rejected"]);

/**
 * Response of POST /api/versions/:id/findings -> what the "Playbook review" section shows.
 * Only the counterparty's own changes (tracked / compare); untracked ones have their own section.
 * Counter-proposals follow the Word flow: reject their change, then suggest ours, as the
 * smallest edit from our text (what the paragraph says once theirs is rejected).
 */
export function reviewCards(res) {
  const units = (res?.units || []).filter((u) => u.source !== "untracked");
  const cards = units
    .map((u, i) => {
      const f = u.finding || {};
      const verdict = f.verdict || "human_review";
      const cp = f.counterProposal && norm(f.counterProposal) !== norm(u.before) ? norm(f.counterProposal) : null;
      const steps = [];
      let edit = null;
      if (u.state === "open") {
        if (verdict === "accept") steps.push("Accept their suggestion (✓ in the margin).");
        else if (verdict === "human_review") steps.push("Decide yourself: accept (✓) or reject (✗) in the margin.");
        else if (cp) {
          steps.push("Reject their suggestion (✗ in the margin).");
          edit = suggestionEdit(u.before, cp);
          steps.push(edit.kind === "whole" || edit.kind === "none"
            ? "Then paste our wording in Suggesting mode."
            : `Then, in Suggesting mode: ${edit.hint}.`);
        } else if (verdict === "reject") steps.push("Reject their suggestion (✗ in the margin).");
        else steps.push("Reject their suggestion (✗), or keep it open and offer our fallback position in a comment.");
      }
      return {
        id: u.id, order: i, verdict, risk: f.risk || null, ruleIds: f.ruleIds || [],
        title: `${u.section || "Paragraph"} · ${u.kind}`,
        rationale: f.rationale || "", flags: f.flags || [],
        diff: wordDiff(u.before, u.after),
        state: u.state, stateText: STATE_TEXT[u.state] || u.state, handled: HANDLED.has(u.state),
        steps, edit, ours: cp, comment: f.comment || "",
        goTo: [u.after, u.before].filter(Boolean),
        carriedOver: u.status === "carried_over",
      };
    })
    .sort((a, b) => (VERDICT_ORDER[a.verdict] ?? 9) - (VERDICT_ORDER[b.verdict] ?? 9) || a.order - b.order);

  const verdicts = {};
  for (const c of cards) verdicts[c.verdict] = (verdicts[c.verdict] || 0) + 1;
  const handled = cards.filter((c) => c.handled).length;
  const parts = ["reject", "human_review", "negotiate", "accept"].filter((v) => verdicts[v]).map((v) => `${verdicts[v]} ${v.replace("_", " ")}`);
  const source = res?.mode === "live" ? res.model : "mock reviewer";
  return {
    cards,
    summary: cards.length ? `${cards.length} change${cards.length > 1 ? "s" : ""} · ${parts.join(" · ")}` : "No changes to review.",
    progress: `${handled} of ${cards.length} handled`,
    source: res?.playbook ? `Playbook ${res.playbook}, ${source}${res.reviewedAt ? `, ${String(res.reviewedAt).slice(0, 10)}` : ""}` : null,
    allHandled: cards.length > 0 && handled === cards.length,
  };
}
