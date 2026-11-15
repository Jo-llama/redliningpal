// Parse WordprocessingML (from body.getOoxml() or word/document.xml) into
// paragraph records with "before" (counterparty's baseline) and "after"
// (counterparty's proposal) text reconstructed from <w:ins>/<w:del> markup.
//
// Pure module: works in the browser (native DOMParser) and in Node tests
// (pass a DOMParser implementation such as @xmldom/xmldom).

export const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

function isInside(node, localName, stopAt) {
  let p = node.parentNode;
  while (p && p !== stopAt) {
    if (p.namespaceURI === W_NS && p.localName === localName) return p;
    p = p.parentNode;
  }
  return null;
}

/**
 * Walk one <w:p> and rebuild its text before and after the tracked changes.
 * - text in <w:t> outside <w:del>            -> after (and before, unless inside <w:ins>)
 * - text in <w:delText> (inside <w:del>)     -> before only
 * - <w:tab/> / <w:br/>                       -> whitespace
 */
export function readParagraph(p) {
  let before = "";
  let after = "";
  const authors = new Set();
  let insertions = 0;
  let deletions = 0;

  const walk = (node) => {
    for (let c = node.firstChild; c; c = c.nextSibling) {
      if (c.nodeType !== 1) continue;
      if (c.namespaceURI !== W_NS) { walk(c); continue; }
      const name = c.localName;

      if (name === "ins" || name === "del") {
        const a = c.getAttributeNS(W_NS, "author") || c.getAttribute("w:author");
        if (a) authors.add(a);
        if (name === "ins") insertions++; else deletions++;
        walk(c);
      } else if (name === "t") {
        const txt = c.textContent || "";
        const inIns = isInside(c, "ins", p);
        const inDel = isInside(c, "del", p);
        if (!inDel) after += txt;
        if (!inIns) before += txt;
      } else if (name === "delText") {
        before += c.textContent || "";
      } else if (name === "tab") {
        const inIns = isInside(c, "ins", p);
        const inDel = isInside(c, "del", p);
        if (!inDel) after += "\t";
        if (!inIns) before += "\t";
      } else if (name === "br" || name === "cr") {
        const inIns = isInside(c, "ins", p);
        const inDel = isInside(c, "del", p);
        if (!inDel) after += "\n";
        if (!inIns) before += "\n";
      } else if (name === "pPr" || name === "rPr" || name === "txbxContent") {
        // properties carry no text; text boxes are skipped (see README limits)
        continue;
      } else {
        walk(c);
      }
    }
  };
  walk(p);

  return {
    before,
    after,
    authors: [...authors],
    insertions,
    deletions,
    hasChanges: insertions + deletions > 0,
  };
}

/**
 * Return body-level paragraphs in document order (including those in tables),
 * skipping paragraphs nested in text boxes so indexes line up with
 * Word's body.paragraphs collection.
 */
export function parseParagraphs(xmlString, DOMParserImpl) {
  const Parser = DOMParserImpl || globalThis.DOMParser;
  const doc = new Parser().parseFromString(xmlString, "application/xml");
  const bodies = doc.getElementsByTagNameNS(W_NS, "body");
  if (!bodies.length) throw new Error("No <w:body> found in OOXML");
  const body = bodies[0];

  const all = body.getElementsByTagNameNS(W_NS, "p");
  const out = [];
  for (let i = 0; i < all.length; i++) {
    const p = all[i];
    if (isInside(p, "txbxContent", body)) continue;
    out.push({ index: out.length, ...readParagraph(p) });
  }
  return out;
}

/** Collapse whitespace so cosmetic differences don't count as changes. */
export function norm(s) {
  return (s || "").replace(/\s+/g, " ").trim();
}

/**
 * Turn parsed paragraphs into review units: one unit per paragraph that
 * carries tracked changes whose text actually differs.
 */
export function unitsFromTrackedChanges(paragraphs) {
  const units = [];
  for (const p of paragraphs) {
    if (!p.hasChanges) continue;
    if (norm(p.before) === norm(p.after)) continue; // formatting-only
    units.push({
      id: `U${units.length + 1}`,
      paragraphIndex: p.index,
      kind: !norm(p.before) ? "inserted" : !norm(p.after) ? "deleted" : "modified",
      before: norm(p.before),
      after: norm(p.after),
      authors: p.authors,
      source: "tracked",
    });
  }
  return units;
}
