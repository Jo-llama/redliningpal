// Which matter does an email + document belong to? Each signal adds points;
// the decision, the score and the evidence are all stored for audit.

import { coverage } from "./matching.js";

const addr = (s) => (s || "").toLowerCase();
const domainOf = (email) => addr(email).split("@")[1] || "";

/**
 * @param msg  { from, subject, inReplyTo, references[] }
 * @param doc  output of readDocx(), or null for mails without a document
 * @param ctx  { store, config }
 */
export function routeMatter(msg, doc, { store, config }) {
  const db = store.db();
  const scores = new Map(db.matters.map((m) => [m.id, { matterId: m.id, score: 0, evidence: [] }]));
  const add = (matterId, pts, why) => {
    const s = scores.get(matterId);
    if (s) { s.score += pts; s.evidence.push(`+${pts} ${why}`); }
  };

  // 1. stamp -> a version we registered
  if (doc?.stamp) {
    const v = db.versions.find((x) => x.stamp === doc.stamp || x.id === doc.stamp);
    if (v) add(v.matterId, 100, `document stamp matches ${v.label}`);
  }

  // 2. thread: In-Reply-To / References point at a message we know
  const threadIds = [msg.inReplyTo, ...(msg.references || [])].filter(Boolean);
  const known = store.messageIds(threadIds).filter((m) => m.matterId);
  for (const id of new Set(known.map((m) => m.matterId))) add(id, 60, "same email thread");

  // 3. matter tag in subject, e.g. "[SPA-Nordlicht]"
  for (const m of db.matters) {
    if ((msg.subject || "").toLowerCase().includes(`[${m.id.toLowerCase()}]`)) add(m.id, 60, `subject tag [${m.id}]`);
  }

  // 4. sender known for the matter
  const from = addr(msg.from);
  for (const m of db.matters) {
    if ((m.counterpartyAddresses || []).map(addr).includes(from)) add(m.id, 30, "known counterparty address");
    else if ((m.counterpartyDomains || []).map(addr).includes(domainOf(from))) add(m.id, 20, `counterparty domain ${domainOf(from)}`);
  }

  // 5. text fingerprint against every version of every matter
  if (doc) {
    for (const m of db.matters) {
      const best = Math.max(0, ...store.versions(m.id).map((v) =>
        Math.max(coverage(doc.beforeTexts, v.afterTexts), coverage(doc.afterTexts, v.afterTexts))));
      if (best >= 0.4) add(m.id, Math.round(50 * best), `text matches existing version (${Math.round(best * 100)}%)`);
    }
  }

  const ranked = [...scores.values()].filter((s) => s.score > 0).sort((a, b) => b.score - a.score);
  const { autoThreshold = 60, minMargin = 25 } = config.routing || {};
  const [top, second] = ranked;
  const auto = !!top && top.score >= autoThreshold && (!second || top.score - second.score >= minMargin);
  return { matterId: auto ? top.matterId : null, auto, candidates: ranked.slice(0, 3) };
}
