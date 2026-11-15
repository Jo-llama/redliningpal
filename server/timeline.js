// Which registered version is the document open in Word, and where does it
// sit in its matter's version graph? Pure functions over store versions.
//
// Identification is by TEXT, not by stamp: the counterparty edits our stamped
// file, so their versions still carry our stamp (see `docStamp` in intake).
// The stamp only breaks ties between equally good text matches.

import { sameText } from "./matching.js";

export const PROBABLE_THRESHOLD = 0.6;
const AMBIGUOUS_MARGIN = 0.02;

const nonEmpty = (arr = []) => arr.filter(Boolean);
const paraSet = (v) => new Set([...nonEmpty(v.afterTexts), ...nonEmpty(v.beforeTexts)]);
const round = (x) => Math.round(x * 100) / 100;

/**
 * How well a version explains the open document, over paragraph sets
 * (accept-all + reject-all text). Accepting a change keeps its "after"
 * paragraph and rejecting keeps its "before" one; both are in the version's
 * set, so a partly reviewed document still scores high.
 *  - coverage: share of the document's paragraphs the version can produce.
 *    Primary score: it does not drop when every change has been accepted.
 *  - jaccard: symmetric overlap. Tie-breaker: an older version whose text is
 *    contained in a later markup's "before" side covers as well, but less snugly.
 */
export function overlap(doc, version) {
  const a = paraSet(doc), b = paraSet(version);
  if (!a.size || !b.size) return { coverage: 0, jaccard: 0 };
  let both = 0;
  for (const p of a) if (b.has(p)) both++;
  return { coverage: both / a.size, jaccard: both / (a.size + b.size - both) };
}

const stamped = (doc, v) => !!doc.stamp && (doc.stamp === v.stamp || doc.stamp === v.id);

/**
 * @param doc       { beforeTexts, afterTexts, stamp } of the open document (normalised)
 * @param versions  all registered versions (any matter)
 * @returns { status: "exact"|"probable"|"unknown", versionId, matterId, score, rendition, candidates, flags }
 *          `candidates` = best score per matter (for triage), never contract text
 */
export function identifyDocument(doc, versions) {
  const flags = [];
  const byId = new Map(versions.map((v) => [v.id, v]));
  const resolve = (v) => (v.renditionOf && byId.get(v.renditionOf)) || v;

  const scored = versions
    .map((v) => {
      const o = overlap(doc, v);
      return {
        v,
        exact: sameText(v.afterTexts || [], doc.afterTexts) && sameText(v.beforeTexts || [], doc.beforeTexts),
        score: round(o.coverage),
        snug: round(o.jaccard),
        stamped: stamped(doc, v),
      };
    })
    .sort((x, y) =>
      Number(y.exact) - Number(x.exact) || y.score - x.score || y.snug - x.snug ||
      Number(y.stamped) - Number(x.stamped) || String(y.v.receivedAt).localeCompare(String(x.v.receivedAt)));

  const perMatter = new Map();
  for (const s of scored) if ((s.exact || s.score > 0) && !perMatter.has(s.v.matterId)) perMatter.set(s.v.matterId, s);
  const candidates = [...perMatter.values()].slice(0, 3).map((s) => {
    const r = resolve(s.v);
    return { matterId: s.v.matterId, versionId: r.id, label: r.label, score: s.score };
  });

  const top = scored[0];
  if (!top || (!top.exact && top.score < PROBABLE_THRESHOLD)) {
    return { status: "unknown", versionId: null, matterId: null, score: top?.score ?? 0, rendition: false, candidates, flags };
  }

  const target = resolve(top.v);
  const runnerUp = scored.find((s) => resolve(s.v).id !== target.id);
  if (!top.exact && runnerUp && top.score - runnerUp.score <= AMBIGUOUS_MARGIN && top.snug - runnerUp.snug <= AMBIGUOUS_MARGIN) {
    flags.push(`Close match with ${resolve(runnerUp.v).label} (${runnerUp.score} vs ${top.score}); please confirm the version.`);
  }
  if (doc.stamp && !top.stamped && target.side === "ours") {
    flags.push("The document's stamp points to a different version than its text.");
  }
  return {
    status: top.exact ? "exact" : "probable",
    versionId: target.id, matterId: target.matterId, score: top.exact ? 1 : top.score,
    rendition: target !== top.v, candidates, flags,
  };
}

// ---------------------------------------------------------------------------

/** Fields the task pane may see. Whitelist: never texts, never findings. */
function node(v) {
  return {
    id: v.id, label: v.label, side: v.side, status: v.status, receivedAt: v.receivedAt,
    from: v.from ?? null, baseId: v.baseId ?? null, supersedesId: v.supersedesId ?? null,
    supersededBy: v.supersededBy ?? null, hasTrackedChanges: !!v.hasTrackedChanges,
  };
}

/** Follow one link (`supersedesId` or `supersededBy`) inside the matter, stopping on loops. */
function walk(start, field, inMatter, flags) {
  const out = [];
  const seen = new Set([start.id]);
  let cur = start;
  while (cur[field]) {
    const next = inMatter.get(cur[field]);
    if (!next) { flags.push(`${cur.label}: ${field} points outside this matter or to an unknown version.`); break; }
    if (seen.has(next.id)) { flags.push(`Loop in ${field} links at ${next.label}; chain cut.`); break; }
    seen.add(next.id);
    out.push(next);
    cur = next;
  }
  return out;
}

/**
 * Timeline around one version: its base, what it supersedes / is superseded by,
 * whether a newer round exists, untracked-change count, and the matter's chain.
 * @returns null if `currentId` is unknown
 */
export function buildTimeline(versions, currentId) {
  const found = versions.find((v) => v.id === currentId);
  if (!found) return null;
  const flags = [];
  const inMatter = new Map(versions.filter((v) => v.matterId === found.matterId).map((v) => [v.id, v]));
  const current = (found.renditionOf && inMatter.get(found.renditionOf)) || found;

  let base = null;
  if (current.baseId) {
    base = inMatter.get(current.baseId) || null;
    if (!base) flags.push(`Base of ${current.label} is not in this matter.`);
  }
  const supersedes = walk(current, "supersedesId", inMatter, flags);
  const supersededBy = walk(current, "supersededBy", inMatter, flags);

  const rounds = [...inMatter.values()]
    .filter((v) => !v.renditionOf)
    .sort((a, b) => String(a.receivedAt).localeCompare(String(b.receivedAt)) || a.label.localeCompare(b.label));
  const latest = rounds[rounds.length - 1];
  const integrity = current.analysis?.integrity || { checked: false, untracked: 0 };

  return {
    matterId: current.matterId,
    current: node(current),
    openedRendition: found !== current,
    base: base && node(base),
    supersedes: supersedes.map(node),
    supersededBy: supersededBy.map(node),
    newerExists: latest.id !== current.id,
    latest: node(latest),
    untracked: { checked: !!integrity.checked, count: integrity.untracked || 0, acknowledgedAt: current.untrackedAck?.at ?? null },
    chain: rounds.map((v) => ({
      ...node(v),
      role: v.id === current.id ? "current" : v.id === base?.id ? "base" : supersedes.some((s) => s.id === v.id) ? "superseded" : null,
      renditions: [...inMatter.values()].filter((r) => r.renditionOf === v.id).length,
    })),
    flags,
  };
}
