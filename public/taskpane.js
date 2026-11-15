/* global Office, Word, JSZip */
import { parseParagraphs, unitsFromTrackedChanges, norm } from "./lib/ooxml.js";
import { diffParagraphs, attachSections, wordDiff, minimalEdit } from "./lib/diff.js";

const $ = (id) => document.getElementById(id);
const state = { units: [], findings: [], playbook: null, source: "tracked", api16: false, playbooks: [] };

Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Word) return setStatus("Open this add-in in Word.", true);
  state.api16 = Office.context.requirements.isSetSupported("WordApi", "1.6");
  wireUi();
  await loadPlaybooks();
  if (!state.api16) setStatus("This Word build lacks WordApi 1.6: accept/reject of tracked changes is disabled; compare mode still works.");
});

function wireUi() {
  document.querySelectorAll('input[name="source"]').forEach((r) =>
    r.addEventListener("change", () => {
      state.source = r.value;
      if (r.value === "compare") $("baseline").click();
    }),
  );
  $("baseline").addEventListener("change", () => {
    const f = $("baseline").files[0];
    setStatus(f ? `Baseline: ${f.name}` : "No baseline chosen.");
  });
  $("scan").addEventListener("click", () => guard(scanAndReview));
  $("accept-ok").addEventListener("click", () => guard(() => bulk("accept", (f) => f.verdict === "accept")));
  $("comment-all").addEventListener("click", () =>
    guard(() => bulk("comment", (f) => ["reject", "negotiate", "human_review"].includes(f.verdict) && f.comment)),
  );
}

async function loadPlaybooks() {
  const r = await fetch("/api/playbooks").then((x) => x.json());
  state.playbooks = r.playbooks;
  $("playbook").innerHTML = r.playbooks
    .map((p) => `<option value="${p.id}">${esc(p.contractType)} (v${esc(p.version)}, ${p.rules} rules)</option>`)
    .join("");
  const badge = $("mode-badge");
  badge.textContent = r.mode === "live" ? r.model : "mock reviewer";
  badge.className = `badge ${r.mode}`;
}

// ---------- scanning ----------

async function readDocumentParagraphs(context) {
  const ox = context.document.body.getOoxml();
  const paras = context.document.body.paragraphs;
  paras.load("text");
  await context.sync();
  const parsed = parseParagraphs(ox.value);
  return { parsed, wordParas: paras.items };
}

async function scanAndReview() {
  setStatus("Reading document…");
  const { parsed, wordCount } = await Word.run(async (context) => {
    const { parsed, wordParas } = await readDocumentParagraphs(context);
    return { parsed, wordCount: wordParas.length };
  });
  if (Math.abs(parsed.length - wordCount) > 1) {
    setStatus(`Warning: ${parsed.length} parsed vs ${wordCount} Word paragraphs (text boxes or content controls?). Navigation may be off.`);
  }

  let units;
  if (state.source === "tracked") {
    units = unitsFromTrackedChanges(parsed);
    if (!units.length) return setStatus("No tracked changes found. Did the counterparty send a clean copy? Try compare mode.");
  } else {
    const file = $("baseline").files[0];
    if (!file) { $("baseline").click(); return setStatus("Choose your last version (.docx) first."); }
    const zip = await JSZip.loadAsync(await file.arrayBuffer());
    const xml = await zip.file("word/document.xml").async("string");
    const oldParas = parseParagraphs(xml).map((p) => p.after);
    units = diffParagraphs(oldParas, parsed.map((p) => p.after));
    if (!units.length) return setStatus("No differences against your last version.");
  }
  units = attachSections(units, parsed.map((p) => p.after)).map((u) => ({
    ...u,
    expect: { before: u.source === "tracked" ? u.before : u.after, after: u.after },
    done: false,
  }));

  setStatus(`Reviewing ${units.length} change(s) against the playbook…`);
  const playbookId = $("playbook").value;
  const resp = await fetch("/api/review", {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ playbookId, units: units.map(({ id, section, kind, before, after }) => ({ id, section, kind, before, after })) }),
  });
  const data = await resp.json();
  if (!resp.ok) throw new Error(data.error || `review failed (${resp.status})`);

  state.units = units;
  state.findings = data.findings;
  state.playbook = await fetch(`/playbooks/${playbookId}.json`).then((x) => (x.ok ? x.json() : null)).catch(() => null);
  render(data.issues || []);
  setStatus(`${units.length} change(s) reviewed (${data.mode}${data.model ? ": " + data.model : ""}).`);
}

// ---------- rendering ----------

const LABEL = { accept: "Accept", negotiate: "Negotiate", reject: "Reject", human_review: "Human review" };
const COLOR = { accept: "var(--accept)", negotiate: "var(--negotiate)", reject: "var(--reject)", human_review: "var(--review)" };

function render(issues) {
  $("summary").hidden = false;
  const counts = {};
  state.findings.forEach((f) => (counts[f.verdict] = (counts[f.verdict] || 0) + 1));
  $("counts").innerHTML = Object.keys(LABEL)
    .filter((k) => counts[k])
    .map((k) => `<span style="background:${COLOR[k]}">${counts[k]} ${LABEL[k]}</span>`)
    .join("");
  $("issues-box").hidden = !issues.length;
  $("issues-count").textContent = issues.length;
  $("issues").innerHTML = issues.map((i) => `<li>${esc(i)}</li>`).join("");

  const cards = $("cards");
  cards.innerHTML = "";
  const order = { reject: 0, human_review: 1, negotiate: 2, accept: 3 };
  const rows = state.units.map((u, k) => ({ u, f: state.findings[k] }))
    .sort((a, b) => order[a.f.verdict] - order[b.f.verdict] || a.u.paragraphIndex - b.u.paragraphIndex);

  for (const { u, f } of rows) {
    const node = $("card-tpl").content.firstElementChild.cloneNode(true);
    node.classList.add(f.verdict);
    node.dataset.unit = u.id;
    node.querySelector(".verdict").textContent = LABEL[f.verdict];
    node.querySelector(".risk").textContent = `${f.risk} risk · ${u.kind}`;
    node.querySelector(".section").textContent = u.section || "";
    node.querySelector(".diff").innerHTML = wordDiff(u.before, u.after)
      .map((s) => (s.op === "=" ? esc(s.text) : s.op === "+" ? `<ins>${esc(s.text)}</ins>` : `<del>${esc(s.text)}</del>`))
      .join("");
    node.querySelector(".rationale").textContent = f.rationale;
    node.querySelector(".rules").textContent = f.ruleIds.length
      ? "Playbook: " + f.ruleIds.map((id) => ruleLabel(id)).join("; ")
      : "Playbook: no rule cited";
    node.querySelector(".flags").innerHTML = f.flags.map((x) => `<li>${esc(x)}</li>`).join("");
    node.querySelector(".counter").value = f.counterProposal || "";
    node.querySelector(".comment").value = f.comment || "";

    const trackedOnly = u.source === "tracked" && !state.api16;
    node.querySelector('[data-act="accept"]').disabled = trackedOnly;
    node.querySelector('[data-act="reject"]').disabled = trackedOnly;
    node.querySelectorAll(".actions button").forEach((b) =>
      b.addEventListener("click", () => guard(() => act(b.dataset.act, u, f, node))),
    );
    cards.appendChild(node);
  }
}

function ruleLabel(id) {
  const r = state.playbook?.rules?.find((x) => x.id === id);
  return r ? `${id} ${r.topic}` : id;
}

function markDone(node, u, label) {
  u.done = true;
  node.classList.add("done");
  const tag = node.querySelector(".done-tag");
  tag.hidden = false;
  tag.textContent = label;
}

// ---------- document actions ----------

/**
 * Find the Word paragraph for a unit. Paragraph indexes shift when whole
 * paragraphs are accepted or rejected, so match on the text we expect the
 * paragraph to have now, preferring the candidate closest to the last index.
 */
async function locate(context, u) {
  const { parsed, wordParas } = await readDocumentParagraphs(context);
  const want = u.expect;
  if (!want.after && !want.before) {
    // a paragraph the counterparty removed in a clean copy has no text to match
    return wordParas[Math.min(u.paragraphIndex, wordParas.length - 1)];
  }
  const score = (p) =>
    want.after ? norm(p.after) === want.after : !norm(p.after) && norm(p.before) === want.before;
  const hits = parsed.filter(score).map((p) => p.index);
  if (!hits.length) {
    // deleted paragraphs in compare mode, or text edited by hand: fall back to last index
    return wordParas[Math.min(u.paragraphIndex, wordParas.length - 1)];
  }
  const idx = hits.reduce((best, i) => (Math.abs(i - u.paragraphIndex) < Math.abs(best - u.paragraphIndex) ? i : best));
  u.paragraphIndex = idx;
  return wordParas[idx];
}

async function act(kind, u, f, node) {
  const counterText = norm(node.querySelector(".counter").value);
  const commentText = node.querySelector(".comment").value.trim();

  await Word.run(async (context) => {
    const para = await locate(context, u);

    if (kind === "goto") {
      para.select();
      await context.sync();
      return;
    }

    if (kind === "comment") {
      if (!commentText) throw new Error("Comment is empty.");
      para.getRange("Whole").insertComment(commentText);
      await context.sync();
      node.querySelector('[data-act="comment"]').textContent = "Commented ✓";
      return;
    }

    if (kind === "accept") {
      if (u.source === "tracked") {
        para.getTrackedChanges().acceptAll();
        await context.sync();
      }
      u.expect = { before: u.after, after: u.after };
      markDone(node, u, "accepted");
      return;
    }

    if (kind === "reject") {
      if (u.source === "tracked") {
        para.getTrackedChanges().rejectAll();
        await context.sync();
      } else {
        await trackedReplace(context, para, u, u.after, u.before);
      }
      u.expect = { before: u.before, after: u.before };
      markDone(node, u, "rejected");
      return;
    }

    if (kind === "counter") {
      if (!counterText) throw new Error("Write a counter-proposal first (or use Reject).");
      let current = u.after;
      if (u.source === "tracked" && u.before) {
        // restore our text first, then write the counter-proposal as our own change.
        // (A paragraph the counterparty inserted is edited in place instead.)
        if (!state.api16) throw new Error("Needs WordApi 1.6 to reject the counterparty's markup first.");
        para.getTrackedChanges().rejectAll();
        await context.sync();
        current = u.before;
      }
      await trackedReplace(context, para, u, current, counterText);
      if (commentText) {
        para.getRange("Whole").insertComment(commentText);
        await context.sync();
      }
      u.expect = { before: u.before, after: counterText };
      markDone(node, u, "countered");
    }
  });
}

/**
 * Write `newText` over `oldText` in `para` as OUR tracked change, touching
 * only the words that differ. Falls back to replacing the paragraph text.
 */
async function trackedReplace(context, para, u, oldText, newText) {
  const doc = context.document;
  doc.load("changeTrackingMode");
  await context.sync();
  const previous = doc.changeTrackingMode;
  doc.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
  await context.sync();

  try {
    if (u.source === "compare" && u.kind === "deleted" && !norm(oldText)) {
      // counterparty deleted our paragraph: re-insert it before the anchor
      para.insertParagraph(newText, "Before");
      await context.sync();
      return;
    }
    if (!norm(newText)) {
      para.delete();
      await context.sync();
      return;
    }
    const { prefix, oldMiddle, newMiddle } = minimalEdit(norm(oldText), newText);
    if (oldMiddle && oldMiddle.length <= 250) {
      const hits = para.search(searchSafe(oldMiddle), { matchCase: true });
      hits.load("items");
      await context.sync();
      if (hits.items.length === 1) {
        hits.items[0].insertText(newMiddle, "Replace");
        await context.sync();
        return;
      }
    } else if (!oldMiddle && prefix) {
      const anchor = prefix.trimEnd().split(/\s+/).slice(-6).join(" ");
      const hits = para.search(searchSafe(anchor), { matchCase: true });
      hits.load("items");
      await context.sync();
      if (hits.items.length === 1 && anchor.length <= 250) {
        const lead = /\s$/.test(prefix) ? "" : " ";
        hits.items[0].insertText(lead + newMiddle.trimStart(), "After");
        await context.sync();
        return;
      }
    }
    // fallback: replace the paragraph's text (formatting inside it is lost)
    para.insertText(newText, "Replace");
    await context.sync();
  } finally {
    doc.changeTrackingMode = previous;
    await context.sync();
  }
}

/** Word's search treats ^ as a control prefix even without wildcards. */
function searchSafe(s) {
  return s.replace(/\^/g, "^^");
}

async function bulk(kind, predicate) {
  let n = 0;
  for (let k = 0; k < state.units.length; k++) {
    const u = state.units[k], f = state.findings[k];
    if (u.done || !predicate(f)) continue;
    const node = document.querySelector(`.card[data-unit="${u.id}"]`);
    await act(kind, u, f, node);
    n++;
  }
  setStatus(`${kind === "accept" ? "Accepted" : "Commented on"} ${n} change(s).`);
}

// ---------- helpers ----------

async function guard(fn) {
  document.querySelectorAll("button").forEach((b) => (b.dataset.wasDisabled = b.disabled, (b.disabled = true)));
  try {
    await fn();
  } catch (e) {
    console.error(e);
    setStatus(e.message || String(e), true);
  } finally {
    document.querySelectorAll("button").forEach((b) => (b.disabled = b.dataset.wasDisabled === "true"));
  }
}

function setStatus(msg, error = false) {
  const s = $("status");
  s.textContent = msg;
  s.classList.toggle("error", error);
}

function esc(s) {
  return String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]);
}
