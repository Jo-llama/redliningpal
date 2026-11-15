// Email intake: raw RFC 822 message in -> versions registered, compared
// against the right base, reviewed, and a short report out.

import crypto from "node:crypto";
import { simpleParser } from "mailparser";
import { readDocx } from "./docx.js";
import { rankBases, sameText } from "./matching.js";
import { routeMatter } from "./routing.js";
import { classifyIntent } from "./intent.js";
import { reviewUnits, loadPlaybook } from "./review-service.js";
import { unitsFromTrackedChanges } from "../public/lib/ooxml.js";
import { diffParagraphs, attachSections } from "../public/lib/diff.js";

const DOCX_RE = /\.docx$/i;
const FWD_RE = /^\s*(fwd?|wg|tr|pd|fw)\s*:/i;

/**
 * @param raw   Buffer with the full email (RFC 822)
 * @param ctx   { store, config, now?, forceMatterId?, notify? }
 */
export async function processMessage(raw, ctx) {
  const { store, config } = ctx;
  const mail = await simpleParser(raw);
  const messageId = mail.messageId || `<generated-${crypto.randomUUID()}>`;
  const existing = store.message(messageId);
  if (existing && !(ctx.forceMatterId && existing.outcome === "triage")) {
    return { messageId, outcome: "duplicate", note: `already processed (${existing.outcome})` };
  }
  const emlHash = store.putBlob("mail", raw, ".eml");

  // --- who sent it, and in which direction ---------------------------------
  const ours = (config.ourAddresses || []).map((a) => a.toLowerCase());
  let from = (mail.from?.value?.[0]?.address || "").toLowerCase();
  let direction = ours.includes(from) ? "outbound" : "inbound";
  const forwarded = direction === "outbound" && FWD_RE.test(mail.subject || "");
  if (forwarded) {
    // a lawyer forwarded the counterparty's mail to the intake address
    // (or forwarded one of our own sent mails - then it stays outbound)
    const m = (mail.text || "").match(/^\s*(?:From|Von|Od)\s*:.*?([\w.+-]+@[\w-]+(?:\.[\w-]+)+)/im);
    const original = m ? m[1].toLowerCase() : "";
    direction = ours.includes(original) ? "outbound" : "inbound";
    from = original || from;
  }
  const msg = {
    messageId, from, direction, forwarded,
    subject: mail.subject || "",
    date: (mail.date || ctx.now || new Date()).toISOString(),
    inReplyTo: mail.inReplyTo || null,
    references: [].concat(mail.references || []),
  };

  // --- attachments ------------------------------------------------------------
  const skipped = [];
  const docs = [];
  for (const a of mail.attachments || []) {
    const name = a.filename || "(unnamed)";
    if (a.contentDisposition === "inline" && /^image\//.test(a.contentType)) continue;
    if (DOCX_RE.test(name) || a.contentType === "application/vnd.openxmlformats-officedocument.wordprocessingml.document") {
      docs.push({ name, buf: a.content });
    } else if (/\.docm$/i.test(name)) {
      skipped.push({ name, reason: "macro-enabled document rejected" });
    } else if (/\.pdf$/i.test(name)) {
      skipped.push({ name, reason: "PDF cannot be redlined - ask for the Word file if no .docx came with it" });
    } else if (/\.doc$/i.test(name)) {
      skipped.push({ name, reason: "legacy .doc - convert to .docx and re-submit" });
    } else {
      skipped.push({ name, reason: "not a contract document" });
    }
  }

  const intent = classifyIntent({ subject: msg.subject, text: mail.text }, docs.length > 0);
  const report = { messageId, direction, from, subject: msg.subject, intent: intent.intent, skipped, documents: [] };

  if (!docs.length) {
    const routing = routeMatter(msg, null, ctx);
    store.addMessage({ ...msg, emlHash, matterId: routing.matterId, intent, outcome: "no_document" });
    return { ...report, outcome: "no_document", matterId: routing.matterId };
  }

  // parse every document first; routing uses the strongest one
  for (const d of docs) {
    d.hash = store.putBlob("files", d.buf, ".docx");
    d.known = store.versionByHash(d.hash);
    if (!d.known) d.doc = await readDocx(d.buf);
  }
  const lead = docs.find((d) => d.doc?.hasTrackedChanges) || docs.find((d) => d.doc) || docs[0];
  const routing = ctx.forceMatterId
    ? { matterId: ctx.forceMatterId, auto: false, candidates: [{ matterId: ctx.forceMatterId, score: 0, evidence: ["assigned by user"] }] }
    : routeMatter(msg, lead.doc || null, ctx);
  report.routing = routing;

  if (!routing.matterId && docs.some((d) => !d.known)) {
    const t = store.addTriage({
      id: `T-${crypto.randomUUID().slice(0, 8)}`, status: "open", messageId, emlHash, subject: msg.subject, from,
      candidates: routing.candidates, createdAt: new Date().toISOString(),
    });
    store.addMessage({ ...msg, emlHash, matterId: null, intent, outcome: "triage", triageId: t.id });
    report.outcome = "triage";
    report.triageId = t.id;
    notify(ctx, `Needs a matter: "${msg.subject}" from ${from || "unknown"} (triage ${t.id})`);
    return report;
  }

  const matter = store.matter(routing.matterId);
  for (const d of docs) {
    if (d.known) {
      report.documents.push({ name: d.name, outcome: "known_file", versionLabel: d.known.label });
      continue;
    }
    const res = direction === "outbound"
      ? registerOurs(d, msg, matter, store)
      : await registerTheirs(d, msg, matter, intent, ctx);
    report.documents.push(res);
  }

  store.addMessage({ ...msg, emlHash, matterId: matter.id, intent, outcome: "processed" });
  report.outcome = "processed";
  report.matterId = matter.id;
  for (const d of report.documents.filter((x) => x.summary)) notify(ctx, `[${matter.id}] ${d.summary}`);
  return report;
}

// ---------------------------------------------------------------------------

function nextLabel(store, matterId, supersedes) {
  const vs = store.versions(matterId);
  if (supersedes) {
    const root = supersedes.label.split(".")[0];
    const n = vs.filter((v) => v.label.startsWith(root + ".")).length + 1;
    return `${root}.${n}`;
  }
  return `v${vs.filter((v) => !v.supersedesId && !v.renditionOf).length + 1}`;
}

function newerThan(store, matterId, version, side) {
  return store.versions(matterId).some((v) => v.side === side && v.receivedAt > version.receivedAt && !v.renditionOf);
}

function registerOurs(d, msg, matter, store) {
  const flags = [];
  const doc = d.doc;
  const ourVs = store.versions(matter.id).filter((v) => v.side === "ours" && v.status === "active");
  const theirVs = store.versions(matter.id).filter((v) => v.side === "theirs" && v.status === "active");

  // our version responds to their latest active version (verified by text)
  const base = theirVs.length ? rankBases(doc, theirVs).best : null;

  // our own correction: same base, they have not replied yet
  const prev = ourVs.filter((v) => v.baseId === (base?.versionId || null)).sort((a, b) => b.receivedAt.localeCompare(a.receivedAt))[0];
  const supersedes = prev && !newerThan(store, matter.id, prev, "theirs") ? prev : null;

  let id = crypto.randomUUID();
  let stamp = null;
  if (doc.stamp && !store.db().versions.some((v) => v.stamp === doc.stamp)) {
    id = stamp = doc.stamp; // stamped before sending: ideal case
  } else {
    flags.push(doc.stamp
      ? "Sent with a stamp of an earlier version - base detection for the reply will rely on text matching. Stamp before sending (npm run stamp)."
      : "Sent without a stamp - base detection for the reply will rely on text matching.");
  }

  const v = store.addVersion({
    id, stamp, matterId: matter.id, side: "ours", label: nextLabel(store, matter.id, supersedes),
    baseId: base?.versionId || null, supersedesId: supersedes?.id || null, status: "active",
    receivedAt: msg.date, messageId: msg.messageId, from: msg.from, fileName: d.name, fileHash: d.hash,
    hasTrackedChanges: doc.hasTrackedChanges, afterTexts: doc.afterTexts, beforeTexts: doc.beforeTexts, flags,
  });
  if (supersedes) store.updateVersion(supersedes.id, { status: "superseded", supersededBy: v.id });
  return {
    name: d.name, outcome: "registered", side: "ours", versionLabel: v.label,
    baseLabel: base?.label || null, supersedes: supersedes?.label || null, flags,
    summary: `Our ${v.label} registered${supersedes ? ` (replaces ${supersedes.label})` : ""}.`,
  };
}

async function registerTheirs(d, msg, matter, intent, ctx) {
  const { store } = ctx;
  const doc = d.doc;
  const flags = [];
  const all = store.versions(matter.id);

  // 1. clean copy of something we already have? -> rendition, not a new version
  const twin = all.find((v) => v.side === "theirs" && sameText(v.afterTexts, doc.afterTexts));
  if (twin && !doc.hasTrackedChanges) {
    store.addVersion({
      id: crypto.randomUUID(), matterId: matter.id, side: "theirs", label: `${twin.label} (clean)`,
      renditionOf: twin.id, baseId: twin.baseId, status: "rendition", receivedAt: msg.date, messageId: msg.messageId,
      from: msg.from, fileName: d.name, fileHash: d.hash, hasTrackedChanges: false, afterTexts: doc.afterTexts, beforeTexts: doc.beforeTexts,
    });
    return { name: d.name, outcome: "rendition", versionLabel: twin.label, summary: `Clean copy of ${twin.label} filed (no new review).` };
  }

  // 2. base = which of OUR versions they worked on
  const ourVs = all.filter((v) => v.side === "ours");
  if (!ourVs.length) flags.push("No version of ours in this matter yet - compared without a base.");
  const { best: base, ranked, flags: baseFlags } = rankBases(doc, ourVs);
  flags.push(...baseFlags);
  const baseV = base ? store.version(base.versionId) : null;

  // 3. correction of their earlier version? same base and we have not answered
  const prev = all
    .filter((v) => v.side === "theirs" && v.status === "active" && v.baseId === (baseV?.id || null))
    .sort((a, b) => b.receivedAt.localeCompare(a.receivedAt))[0];
  const supersedes = prev && !newerThan(store, matter.id, prev, "ours") ? prev : null;
  if (supersedes && intent.intent !== "corrected_version") {
    flags.push(`Same base as ${supersedes.label} and not announced as a correction - treated as replacing ${supersedes.label}; please confirm.`);
  }
  if (!supersedes && intent.intent === "corrected_version") {
    flags.push("Email says this is a corrected version, but no earlier version on the same base was found.");
  }

  const v = store.addVersion({
    id: crypto.randomUUID(), stamp: null, matterId: matter.id, side: "theirs",
    label: nextLabel(store, matter.id, supersedes), baseId: baseV?.id || null, baseCandidates: ranked.slice(0, 3),
    supersedesId: supersedes?.id || null, status: "active", receivedAt: msg.date, messageId: msg.messageId,
    from: msg.from, fileName: d.name, fileHash: d.hash, hasTrackedChanges: doc.hasTrackedChanges,
    afterTexts: doc.afterTexts, beforeTexts: doc.beforeTexts, docStamp: doc.stamp, intent: intent.intent, flags,
  });
  if (supersedes) store.updateVersion(supersedes.id, { status: "superseded", supersededBy: v.id });

  // 4. analysis against the base
  const analysis = await analyse(doc, baseV, supersedes, matter, ctx);
  store.updateVersion(v.id, { analysis });

  const c = analysis.counts;
  const parts = [
    `${v.label} from ${msg.from} (base ${baseV?.label || "none"}${supersedes ? `, replaces ${supersedes.label}` : ""})`,
    `${analysis.units.length} change(s)`,
    analysis.integrity.untracked ? `${analysis.integrity.untracked} UNTRACKED` : null,
    c.reject ? `${c.reject} reject` : null,
    c.human_review ? `${c.human_review} human review` : null,
    supersedes ? `${analysis.delta.newUnits} new since ${supersedes.label}` : null,
  ].filter(Boolean);
  return {
    name: d.name, outcome: "registered", side: "theirs", versionLabel: v.label, versionId: v.id,
    baseLabel: baseV?.label || null, supersedes: supersedes?.label || null, flags, counts: c,
    integrity: analysis.integrity, delta: analysis.delta, summary: parts.join(", ") + ".",
  };
}

async function analyse(doc, baseV, supersedes, matter, ctx) {
  const { store } = ctx;
  const playbook = loadPlaybook(matter.playbook);
  let units = [];
  const integrity = { checked: false, untracked: 0 };

  if (doc.hasTrackedChanges) {
    units = unitsFromTrackedChanges(doc.paragraphs).map((u) => ({ ...u, source: "tracked" }));
    if (baseV) {
      // "reject all" of their file must equal our base; anything else was edited without tracking
      const silent = diffParagraphs(baseV.afterTexts, doc.beforeTexts).map((u) => ({ ...u, source: "untracked" }));
      integrity.checked = true;
      integrity.untracked = silent.length;
      units = units.concat(silent);
    }
  } else if (baseV) {
    units = diffParagraphs(baseV.afterTexts, doc.afterTexts).map((u) => ({ ...u, source: "compare" }));
  }
  units = attachSections(units.sort((a, b) => a.paragraphIndex - b.paragraphIndex), doc.afterTexts)
    .map((u, i) => ({ ...u, id: `U${i + 1}` }));

  // reuse findings for identical changes (e.g. carried over from a superseded version)
  const key = (u) => crypto.createHash("sha256").update(`${playbook.id}@${playbook.version}|${u.source}|${u.before}|${u.after}`).digest("hex");
  const prevKeys = new Set((supersedes?.analysis?.units || []).map((u) => u.cacheKey));
  const todo = [];
  const results = new Map();
  for (const u of units) {
    u.cacheKey = key(u);
    const hit = store.cached(u.cacheKey);
    if (hit) results.set(u.id, { ...hit, reused: true }); else todo.push(u);
  }
  const review = await reviewUnits(playbook, todo, { mock: ctx.mock });
  todo.forEach((u, i) => {
    results.set(u.id, review.findings[i]);
    store.cache(u.cacheKey, review.findings[i]);
  });
  store.save();

  const rank = { accept: 0, negotiate: 1, human_review: 2, reject: 3 };
  const out = units.map((u) => {
    const f = { ...results.get(u.id), flags: [...(results.get(u.id).flags || [])] };
    if (u.source === "untracked") {
      f.flags.unshift("Changed WITHOUT track changes");
      f.risk = "high";
      if (rank[f.verdict] < rank.human_review) f.verdict = "human_review";
    }
    return {
      id: u.id, source: u.source, kind: u.kind, section: u.section, before: u.before, after: u.after,
      cacheKey: u.cacheKey, status: supersedes ? (prevKeys.has(u.cacheKey) ? "carried_over" : "new") : "new", finding: f,
    };
  });

  const counts = {};
  for (const u of out) counts[u.finding.verdict] = (counts[u.finding.verdict] || 0) + 1;
  return {
    reviewedAt: new Date().toISOString(), playbook: `${playbook.id}@${playbook.version}`,
    mode: review.mode, model: review.model, harnessIssues: review.issues,
    integrity, counts, units: out,
    delta: supersedes
      ? { supersededLabel: supersedes.label, newUnits: out.filter((u) => u.status === "new").length, carriedOver: out.filter((u) => u.status === "carried_over").length }
      : null,
  };
}

function notify(ctx, text) {
  if (ctx.notify) ctx.notify(text); else console.log(`[notify] ${text}`);
}
