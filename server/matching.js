// Text fingerprints: which known version does an incoming document come from?

const nonEmpty = (arr) => arr.filter(Boolean);

/** Share of `doc` paragraphs that appear verbatim (normalised) in `candidate`. */
export function coverage(docParas, candidateParas) {
  const doc = nonEmpty(docParas);
  if (!doc.length) return 0;
  const pool = new Map();
  for (const p of nonEmpty(candidateParas)) pool.set(p, (pool.get(p) || 0) + 1);
  let hit = 0;
  for (const p of doc) {
    const c = pool.get(p);
    if (c) { hit++; pool.set(p, c - 1); }
  }
  return hit / doc.length;
}

/** Two renditions of the same text (e.g. clean copy of an existing markup). */
export function sameText(a, b) {
  return nonEmpty(a).join("\n") === nonEmpty(b).join("\n");
}

/**
 * Rank our versions as the base of an incoming document.
 * The incoming "reject all" text should equal its base exactly, so use
 * beforeTexts; for clean copies before == after and the score is lower.
 * A stamp is strong evidence but is verified, never trusted blindly:
 * the file may have been re-used for a later round without re-stamping.
 */
export function rankBases(doc, oursVersions) {
  const ranked = oursVersions
    .map((v) => {
      const cov = coverage(doc.beforeTexts, v.afterTexts);
      const stamped = doc.stamp && (doc.stamp === v.stamp || doc.stamp === v.id);
      return { versionId: v.id, label: v.label, coverage: round(cov), stamped: !!stamped, score: round(cov + (stamped ? 0.15 : 0)) };
    })
    .sort((a, b) => b.score - a.score);
  const best = ranked[0] || null;
  const flags = [];
  if (best && doc.stamp && !best.stamped) {
    const s = ranked.find((r) => r.stamped);
    flags.push(s
      ? `Stamp points to ${s.label}, but the text matches ${best.label} better (${s.coverage} vs ${best.coverage}). Using ${best.label}.`
      : "Stamp not found among our versions in this matter.");
  }
  if (best && best.coverage < 0.6) flags.push(`Weak base match (${best.coverage}); please confirm the base version.`);
  return { best, ranked, flags };
}

const round = (x) => Math.round(x * 100) / 100;
