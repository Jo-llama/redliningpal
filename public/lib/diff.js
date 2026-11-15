// Paragraph-level compare for when the counterparty sends a clean copy
// (no tracked changes). Aligns our last version against the current
// document and emits review units in the same shape as tracked-change units.

import { norm } from "./ooxml.js";

/** Longest common subsequence over normalised paragraph text. */
function lcsPairs(a, b) {
  const n = a.length, m = b.length;
  const dp = Array.from({ length: n + 1 }, () => new Uint32Array(m + 1));
  for (let i = n - 1; i >= 0; i--) {
    for (let j = m - 1; j >= 0; j--) {
      dp[i][j] = a[i] === b[j] ? dp[i + 1][j + 1] + 1 : Math.max(dp[i + 1][j], dp[i][j + 1]);
    }
  }
  const pairs = [];
  let i = 0, j = 0;
  while (i < n && j < m) {
    if (a[i] === b[j]) { pairs.push([i, j]); i++; j++; }
    else if (dp[i + 1][j] >= dp[i][j + 1]) i++;
    else j++;
  }
  return pairs;
}

/** Word-overlap similarity (Dice coefficient) between two paragraphs. */
export function similarity(x, y) {
  const wx = norm(x).toLowerCase().split(" ").filter(Boolean);
  const wy = norm(y).toLowerCase().split(" ").filter(Boolean);
  if (!wx.length && !wy.length) return 1;
  if (!wx.length || !wy.length) return 0;
  const counts = new Map();
  for (const w of wx) counts.set(w, (counts.get(w) || 0) + 1);
  let common = 0;
  for (const w of wy) {
    const c = counts.get(w);
    if (c) { common++; counts.set(w, c - 1); }
  }
  return (2 * common) / (wx.length + wy.length);
}

/**
 * Inside a gap between two LCS anchors, pair old/new paragraphs greedily in
 * order when they are similar enough to be "the same clause, edited";
 * everything else is a pure insertion or deletion.
 */
function alignGap(oldGap, newGap, threshold) {
  const ops = [];
  let i = 0, j = 0;
  while (i < oldGap.length || j < newGap.length) {
    const o = oldGap[i], nw = newGap[j];
    if (o && nw && similarity(o.text, nw.text) >= threshold) {
      ops.push({ type: "modified", old: o, new: nw }); i++; j++;
    } else if (o && nw) {
      // look ahead one step to decide whether o was deleted or nw inserted
      const nextNewMatches = newGap[j + 1] && similarity(o.text, newGap[j + 1].text) >= threshold;
      if (nextNewMatches) { ops.push({ type: "inserted", new: nw }); j++; }
      else { ops.push({ type: "deleted", old: o }); i++; }
    } else if (o) {
      ops.push({ type: "deleted", old: o }); i++;
    } else {
      ops.push({ type: "inserted", new: nw }); j++;
    }
  }
  return ops;
}

/**
 * @param {string[]} oldParas  our last version, paragraph texts
 * @param {string[]} newParas  current document, paragraph texts (index = Word paragraph index)
 * @returns review units; paragraphIndex points into newParas (for deletions:
 *          the next surviving paragraph, so the UI can still navigate there)
 */
export function diffParagraphs(oldParas, newParas, { threshold = 0.5 } = {}) {
  const A = oldParas.map((t, i) => ({ i, text: t, key: norm(t) })).filter((p) => p.key);
  const B = newParas.map((t, i) => ({ i, text: t, key: norm(t) })).filter((p) => p.key);
  const pairs = lcsPairs(A.map((p) => p.key), B.map((p) => p.key));

  const ops = [];
  let ai = 0, bi = 0;
  for (const [pa, pb] of [...pairs, [A.length, B.length]]) {
    const prevNew = bi > 0 ? B[bi - 1].i : -1; // Word index of the anchor before this gap
    for (const op of alignGap(A.slice(ai, pa), B.slice(bi, pb), threshold)) ops.push({ ...op, prevNew });
    ai = pa + 1; bi = pb + 1;
  }

  const units = [];
  let lastNewIndex = -1;
  for (const op of ops) {
    lastNewIndex = Math.max(lastNewIndex, op.prevNew);
    const anchor = op.new ? op.new.i : nextIndexAfter(B, lastNewIndex, newParas.length);
    if (op.new) lastNewIndex = op.new.i;
    units.push({
      id: `U${units.length + 1}`,
      paragraphIndex: anchor,
      kind: op.type,
      before: op.old ? norm(op.old.text) : "",
      after: op.new ? norm(op.new.text) : "",
      authors: [],
      source: "compare",
    });
  }
  return units.sort((x, y) => x.paragraphIndex - y.paragraphIndex).map((u, k) => ({ ...u, id: `U${k + 1}` }));
}

function nextIndexAfter(B, lastIndex, total) {
  const next = B.find((p) => p.i > lastIndex);
  return next ? next.i : Math.max(0, total - 1);
}

/**
 * Word-level diff for display: returns [{op: "=", "-", "+", text}].
 */
export function wordDiff(before, after) {
  const a = (before || "").split(/(\s+)/).filter((t) => t !== "");
  const b = (after || "").split(/(\s+)/).filter((t) => t !== "");
  const n = a.length, m = b.length;
  const dp = Array.from({ length: n + 1 }, () => new Uint32Array(m + 1));
  for (let i = n - 1; i >= 0; i--)
    for (let j = m - 1; j >= 0; j--)
      dp[i][j] = a[i] === b[j] ? dp[i + 1][j + 1] + 1 : Math.max(dp[i + 1][j], dp[i][j + 1]);
  const out = [];
  const push = (op, text) => {
    const last = out[out.length - 1];
    if (last && last.op === op) last.text += text; else out.push({ op, text });
  };
  let i = 0, j = 0;
  while (i < n || j < m) {
    if (i < n && j < m && a[i] === b[j]) { push("=", a[i]); i++; j++; }
    else if (i < n && (j >= m || dp[i + 1][j] >= dp[i][j + 1])) { push("-", a[i]); i++; }
    else { push("+", b[j]); j++; }
  }
  return out;
}

/**
 * Smallest edit that turns `oldText` into `newText`, at word granularity:
 * shared prefix and suffix are kept, only the middle is replaced.
 * Used to write counter-proposals as minimal tracked changes.
 */
export function minimalEdit(oldText, newText) {
  const a = (oldText || "").split(/(\s+)/).filter((t) => t !== "");
  const b = (newText || "").split(/(\s+)/).filter((t) => t !== "");
  let p = 0;
  while (p < a.length && p < b.length && a[p] === b[p]) p++;
  let s = 0;
  while (s < a.length - p && s < b.length - p && a[a.length - 1 - s] === b[b.length - 1 - s]) s++;
  return {
    prefix: a.slice(0, p).join(""),
    oldMiddle: a.slice(p, a.length - s).join(""),
    newMiddle: b.slice(p, b.length - s).join(""),
    suffix: a.slice(a.length - s).join(""),
  };
}

/**
 * Attach the nearest preceding heading to each unit, so the model and the
 * reviewer know which clause a change sits in. A "heading" is a short
 * paragraph that starts with a clause number or is in title-like form.
 */
export function attachSections(units, paragraphTexts) {
  const headingRe = /^((\d+(\.\d+)*\.?)|(§\s*\d+)|(clause\s+\d+))\s+\S/i;
  const headings = paragraphTexts.map((t) => {
    const s = norm(t);
    return s.length > 0 && s.length <= 80 && headingRe.test(s) && !/[.;:,]$/.test(s) ? s : null;
  });
  return units.map((u) => {
    let section = null;
    for (let k = Math.min(u.paragraphIndex, headings.length - 1); k >= 0; k--) {
      if (headings[k]) { section = headings[k]; break; }
    }
    return { ...u, section };
  });
}
