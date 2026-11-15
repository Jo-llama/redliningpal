/* global Office, Word, JSZip */
import { parseParagraphs, unitsFromTrackedChanges, norm } from "./lib/ooxml.js";
import { diffParagraphs, attachSections, wordDiff, minimalEdit } from "./lib/diff.js";
import { versionView, untrackedUnits } from "./lib/version-view.js";

const $ = (id) => document.getElementById(id);
const state = {
  units: [], findings: [], playbook: null, source: "tracked", api16: false, playbooks: [], matters: [],
  identify: null, view: null, triageId: null, opened: {},
};

Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Word) return setStatus("Open this add-in in Word.", true);
  state.api16 = Office.context.requirements.isSetSupported("WordApi", "1.6");
  wireUi();
  await loadPlaybooks();
  state.matters = await fetch("/api/matters").then((x) => (x.ok ? x.json() : [])).catch(() => []);
  await guard(identifyOpenDocument);
  if (!state.api16) setStatus("This Word build lacks WordApi 1.6: accept/reject of tracked changes is disabled; compare mode still works.");
});

function wireUi() {
  document.querySelectorAll('input[name="source"]').forEach((r) =>
    r.addEventListener("change", () => {
      state.source = r.value;
      // with a registered base, compare against it; otherwise ask for a file
      if (r.value === "compare" && !state.view?.base) $("baseline").click();
    }),
  );
  $("recheck").addEventListener("click", () => guard(identifyOpenDocument));
  $("open-base").addEventListener("click", () => guard(openBase));
  $("ack-untracked").addEventListener("change", () => guard(ackUntracked));
  $("triage-file").addEventListener("click", () => guard(fileForTriage));
  $("t-matter").addEventListener("change", () => guard(loadBases));
  document.querySelectorAll('input[name="t-side"]').forEach((r) => r.addEventListener("change", () => guard(loadBases)));
  $("triage-form").addEventListener("submit", (e) => {
    e.preventDefault();
    guard(confirmTriage);
  });
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

// ---------- which version is open (timeline) ----------

async function identifyOpenDocument() {
  $("version-status").hidden = false;
  $("version-status").textContent = "Checking which version this is…";
  const { parsed, stamp } = await Word.run(async (context) => {
    const ox = context.document.body.getOoxml();
    const prop = context.document.properties.customProperties.getItemOrNullObject("rc-ver");
    prop.load("value");
    await context.sync();
    return { parsed: parseParagraphs(ox.value), stamp: prop.isNullObject ? null : String(prop.value) };
  });
  state.opened = { hasTrackedChanges: parsed.some((p) => p.hasChanges) };
  const r = await api("/api/documents/identify", {
    afterTexts: parsed.map((p) => norm(p.after)), beforeTexts: parsed.map((p) => norm(p.before)), stamp,
  });
  if (!r.ok) {
    $("version-status").textContent = "Version check unavailable.";
    throw new Error(r.data.error || `version check failed (${r.status})`);
  }
  showVersion(r.data);
}

function showVersion(identify) {
  state.identify = identify;
  const v = (state.view = versionView(identify, state.opened));
  $("version-status").hidden = true;
  $("version-known").hidden = !v.known;
  $("version-unknown").hidden = v.known;
  if (!v.known) return showUnknown(v);

  $("v-title").textContent = v.title;
  $("v-matter").textContent = v.matterId;
  $("v-match").textContent = v.match;
  $("v-warning").hidden = !v.warning;
  $("v-warning").textContent = v.warning || "";
  $("v-base").textContent = v.base?.text || "none";
  $("open-base").hidden = !v.base;
  $("v-supersedes").textContent = v.supersedes || "nothing";
  $("v-untracked").hidden = !v.untracked.count;
  $("v-untracked-text").textContent = v.untracked.text || "";
  $("ack-untracked").checked = !!v.untracked.acknowledgedAt;
  $("ack-untracked").disabled = !!v.untracked.acknowledgedAt; // an acknowledgement is audited and final
  const ROLE = { base: "base", superseded: "superseded", current: "this document" };
  $("v-chain").innerHTML = v.chain
    .map((c) => `<li class="${c.current ? "current" : ""} ${c.status === "superseded" ? "superseded" : ""}">${esc(c.text)}${c.role ? ` <span class="role">${ROLE[c.role]}</span>` : ""}</li>`)
    .join("");
  $("v-flags").innerHTML = v.flags.map((f) => `<li>${esc(f)}</li>`).join("");

  // the review setup follows the version: mode, compare target, playbook
  const radio = document.querySelector(`input[name="source"][value="${v.source}"]`);
  if (radio) radio.checked = true;
  state.source = v.source;
  $("compare-target").textContent = v.base ? `the base (${identify.timeline.base.label}) from the registry` : "my last version";
  const playbook = state.matters.find((m) => m.id === v.matterId)?.playbook;
  if (playbook && [...$("playbook").options].some((o) => o.value === playbook)) $("playbook").value = playbook;
  applyGate();
}

function showUnknown(v) {
  $("u-candidates").innerHTML = v.candidates.length
    ? v.candidates.map((c) => `<li>${esc(c.text)}</li>`).join("")
    : "<li>No similar document in any matter.</li>";
  $("triage-file").hidden = !!state.triageId;
  $("triage-form").hidden = !state.triageId;
  applyGate();
}

async function openBase() {
  const base = state.view?.base;
  if (!base) return;
  setStatus("Opening the base version…");
  const res = await fetch(`/api/versions/${encodeURIComponent(base.id)}/file`);
  if (!res.ok) throw new Error(`Could not load the base version (${res.status}).`);
  const b64 = toBase64(new Uint8Array(await res.arrayBuffer()));
  await Word.run(async (context) => {
    context.application.createDocument(b64).open();
    await context.sync();
  });
  setStatus(`Opened ${base.text} in a new window.`);
}

async function ackUntracked() {
  const box = $("ack-untracked");
  if (!box.checked || !state.view?.known) return;
  const r = await api(`/api/versions/${encodeURIComponent(state.view.versionId)}/ack-untracked`, {});
  if (!r.ok) {
    box.checked = false;
    throw new Error(r.data.error || `could not record the acknowledgement (${r.status})`);
  }
  state.identify.timeline.untracked.acknowledgedAt = r.data.acknowledgedAt;
  showVersion(state.identify);
  setStatus("Acknowledged. Bulk accept is available again.");
}

/** "Accept all" is blocked while untracked changes are unacknowledged. Run after anything that re-enables buttons. */
function applyGate() {
  const b = $("accept-ok");
  if (state.view?.gateClosed) {
    b.disabled = true;
    b.title = "Blocked: review the changes made without track changes and tick the box at the top first.";
  } else {
    b.disabled = false;
    b.title = "Accept every change the review marked as acceptable";
  }
}

// ---------- unknown document: triage, a lawyer decides ----------

async function fileForTriage() {
  setStatus("Uploading the document for a matter decision…");
  const bytes = await readDocxBytes();
  const name = (Office.context.document.url || "").split(/[\\/]/).pop() || "from-word.docx";
  const r = await api("/api/documents/triage", bytes, { raw: true, headers: { "x-file-name": name } });
  if (r.status === 409) {
    setStatus("This document is already registered.");
    return showVersion({ status: "exact", versionId: r.data.versionId, score: 1, flags: [], timeline: r.data.timeline });
  }
  if (!r.ok) throw new Error(r.data.error || `filing failed (${r.status})`);
  state.triageId = r.data.triageId;
  const likely = new Set(r.data.candidates.map((c) => c.matterId));
  const matters = [...r.data.matters].sort((a, b) => Number(likely.has(b.id)) - Number(likely.has(a.id)));
  $("t-matter").innerHTML = matters.map((m) => `<option value="${esc(m.id)}">${esc(m.id)}: ${esc(m.title || "")}</option>`).join("");
  $("t-id").textContent = state.triageId;
  $("triage-file").hidden = true;
  $("triage-form").hidden = false;
  await loadBases();
  setStatus("Choose the matter, the side and the base, then confirm.");
}

async function loadBases() {
  if (!state.triageId) return;
  const side = document.querySelector('input[name="t-side"]:checked').value;
  const matterId = $("t-matter").value;
  const r = await api(`/api/triage/${encodeURIComponent(state.triageId)}/bases?matterId=${encodeURIComponent(matterId)}&side=${side}`);
  if (!r.ok) throw new Error(r.data.error || `could not load bases (${r.status})`);
  $("t-base").innerHTML = '<option value="">No base: first version in this matter</option>' +
    r.data.bases
      .map((b, i) => `<option value="${esc(b.versionId)}"${i === 0 ? " selected" : ""}>${esc(b.label)}: ${Math.round(b.coverage * 100)}% match${i === 0 ? " (suggested)" : ""}</option>`)
      .join("");
  $("t-flags").innerHTML = r.data.flags.map((f) => `<li>${esc(f)}</li>`).join("");
  $("triage-confirm").textContent = `Register as ${side === "ours" ? "our" : "counterparty"} version in ${matterId}`;
}

async function confirmTriage() {
  const side = document.querySelector('input[name="t-side"]:checked').value;
  const r = await api(`/api/triage/${encodeURIComponent(state.triageId)}/assign`, {
    matterId: $("t-matter").value, side, baseId: $("t-base").value || undefined,
  });
  if (!r.ok) throw new Error(r.data.error || `registration failed (${r.status})`);
  state.triageId = null;
  $("triage-form").hidden = true;
  showVersion({ status: "exact", versionId: r.data.versionId, score: 1, flags: r.data.flags || [], timeline: r.data.timeline });
  setStatus(`Registered as ${r.data.versionLabel} in ${r.data.matterId}.`);
}

/** The open document as .docx bytes (Office "compressed" file, read in slices). */
function readDocxBytes() {
  const ok = Office.AsyncResultStatus.Succeeded;
  return new Promise((resolve, reject) => {
    Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 4 * 1024 * 1024 }, (res) => {
      if (res.status !== ok) return reject(new Error(res.error?.message || "could not read the document"));
      const file = res.value;
      const parts = [];
      const next = (i) => file.getSliceAsync(i, (s) => {
        if (s.status !== ok) { file.closeAsync(); return reject(new Error(s.error?.message || "could not read the document")); }
        parts.push(new Uint8Array(s.value.data));
        if (i + 1 < file.sliceCount) return next(i + 1);
        file.closeAsync();
        const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
        let at = 0;
        for (const p of parts) { out.set(p, at); at += p.length; }
        resolve(out);
      });
      next(0);
    });
  });
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
  let untracked = [];
  if (state.source === "tracked") {
    units = unitsFromTrackedChanges(parsed);
    // edits made without track changes are invisible in the markup; intake found them against the base
    untracked = await loadUntracked(parsed);
    if (!units.length && !untracked.length) return setStatus("No tracked changes found. Did the counterparty send a clean copy? Try compare mode.");
  } else {
    let buf;
    const file = $("baseline").files[0];
    if (file) {
      buf = await file.arrayBuffer();
    } else if (state.view?.base) {
      setStatus(`Comparing with ${state.view.base.text} from the registry…`);
      const res = await fetch(`/api/versions/${encodeURIComponent(state.view.base.id)}/file`);
      if (!res.ok) throw new Error(`Could not load the base version (${res.status}).`);
      buf = await res.arrayBuffer();
    } else {
      $("baseline").click();
      return setStatus("Choose your last version (.docx) first.");
    }
    const zip = await JSZip.loadAsync(buf);
    const xml = await zip.file("word/document.xml").async("string");
    const oldParas = parseParagraphs(xml).map((p) => p.after);
    units = diffParagraphs(oldParas, parsed.map((p) => p.after));
    if (!units.length) return setStatus("No differences against the base version.");
  }
  units = attachSections(units, parsed.map((p) => p.after)).map((u) => ({
    ...u,
    expect: { before: u.source === "tracked" ? u.before : u.after, after: u.after },
    done: false,
  }));

  const playbookId = $("playbook").value;
  let data = { findings: [], issues: [], mode: "registry" };
  if (units.length) {
    setStatus(`Reviewing ${units.length} change(s) against the playbook…`);
    const resp = await fetch("/api/review", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ playbookId, units: units.map(({ id, section, kind, before, after }) => ({ id, section, kind, before, after })) }),
    });
    data = await resp.json();
    if (!resp.ok) throw new Error(data.error || `review failed (${resp.status})`);
  }

  // untracked cards keep the findings intake already validated (always at least human_review)
  state.units = units.concat(untracked);
  state.findings = data.findings.concat(untracked.map((u) => u.finding));
  state.playbook = await fetch(`/playbooks/${playbookId}.json`).then((x) => (x.ok ? x.json() : null)).catch(() => null);
  render(data.issues || []);
  setStatus(`${state.units.length} change(s) reviewed (${data.mode}${data.model ? ": " + data.model : ""})` +
    (untracked.length ? `, incl. ${untracked.length} made without track changes.` : "."));
}

async function loadUntracked(parsed) {
  const v = state.view;
  if (!v?.known || !v.untracked.count) return [];
  const r = await api(`/api/versions/${encodeURIComponent(v.versionId)}/untracked`);
  if (!r.ok) throw new Error(r.data.error || `could not load untracked changes (${r.status})`);
  return untrackedUnits(r.data.units, parsed);
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
    if (u.untracked) node.classList.add("untracked-change");
    node.dataset.unit = u.id;
    node.querySelector(".verdict").textContent = LABEL[f.verdict];
    node.querySelector(".risk").textContent = `${f.risk} risk · ${u.untracked ? "untracked " : ""}${u.kind}`;
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
  if (kind === "accept" && state.view?.gateClosed) {
    throw new Error("Review the changes made without track changes and tick the box at the top first.");
  }
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
    applyGate();
  }
}

/** fetch wrapper: GET without body, POST JSON (or raw bytes with { raw: true }). */
async function api(url, body, { raw = false, headers = {} } = {}) {
  const init = body === undefined ? {} : {
    method: "POST",
    headers: { "content-type": raw ? "application/octet-stream" : "application/json", ...headers },
    body: raw ? body : JSON.stringify(body),
  };
  const res = await fetch(url, init);
  const data = await res.json().catch(() => ({}));
  return { ok: res.ok, status: res.status, data };
}

function toBase64(bytes) {
  let s = "";
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  return btoa(s);
}

function setStatus(msg, error = false) {
  const s = $("status");
  s.textContent = msg;
  s.classList.toggle("error", error);
}

function esc(s) {
  return String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]);
}
