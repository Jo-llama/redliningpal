// gdocs/Sidebar.html in jsdom with a fake google.script.run: what the lawyer sees,
// what each button calls, and that server text is never interpreted as HTML.
import { test } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { JSDOM } from "jsdom";

const SIDEBAR = readFileSync(new URL("../gdocs/Sidebar.html", import.meta.url), "utf8");
const VIEW = readFileSync(new URL("../gdocs/VersionView.html", import.meta.url), "utf8");
const tick = () => new Promise((r) => setTimeout(r, 5));
/** Wait until every fake google.script.run call has been answered (no fixed sleeps: CI load varies). */
async function waitIdle(state) {
  const until = Date.now() + 5000;
  let quiet = 0;
  while (quiet < 3) {
    if (Date.now() > until) throw new Error("sidebar did not settle");
    await tick();
    quiet = state.pending ? 0 : quiet + 1;
  }
}
const XSS = '<img src=x onerror="window.__pwned=1">';

const node = (id, label, side, extra = {}) => ({ id, label, side, status: "active", receivedAt: "2026-11-13T09:00:00Z", hasTrackedChanges: true, ...extra });
const v21 = (untracked = { checked: true, count: 1, acknowledgedAt: null }) => ({
  status: "exact", versionId: "id-v21", score: 1, rendition: false, flags: [], candidates: [],
  opened: { hasTrackedChanges: true, stamped: true },
  timeline: {
    matterId: "SPA-Nordlicht", current: node("id-v21", "v2.1", "theirs"), openedRendition: false,
    base: node("id-v1", "v1", "ours", { receivedAt: "2026-11-10T09:00:00Z" }),
    supersedes: [node("id-v2", "v2", "theirs", { status: "superseded" })], supersededBy: [],
    newerExists: false, latest: node("id-v21", "v2.1", "theirs"), untracked, chain: [], flags: [],
  },
});
const untrackedCards = {
  versionId: "id-v21", checked: true,
  units: [{
    id: "U1", section: "4.4", kind: "modified",
    before: "4.4 None of the limitations apply in case of intent or fraud.",
    after: `${XSS} 4.4 None of the limitations apply in case of intent.`,
    finding: { verdict: "reject", risk: "high", rationale: "<b>red line</b>", counterProposal: null },
  }],
};

/**
 * @param server  { fnName: (...args) => value | throws }  (Code.gs stand-ins)
 * @returns { window, $, calls }
 */
async function open(server) {
  const calls = [];
  const state = { pending: 0 };
  const html = SIDEBAR.replace('<?!= include("VersionView") ?>', VIEW);
  const dom = new JSDOM(html, {
    runScripts: "dangerously",
    beforeParse(window) {
      const runner = (ok, fail) => new Proxy({}, {
        get: (_t, prop) => {
          if (prop === "withSuccessHandler") return (h) => runner(h, fail);
          if (prop === "withFailureHandler") return (h) => runner(ok, h);
          return (...args) => {
            calls.push([prop, ...args]);
            state.pending++;
            setTimeout(() => {
              state.pending--;
              if (!server[prop]) return fail(new Error(`Exception: unexpected call ${String(prop)}`));
              let value;
              try { value = server[prop](...args); } catch (e) { return fail(e); }
              ok(value);
            }, 0);
          };
        },
      });
      window.google = { script: { get run() { return runner(() => {}, () => {}); } } };
    },
  });
  const settle = () => waitIdle(state);
  await settle();
  const $ = (id) => dom.window.document.getElementById(id);
  return { window: dom.window, $, calls, settle };
}
const configured = { getSettingsInfo: () => ({ url: "https://demo.trycloudflare.com", hasToken: true }) };
const visible = (n) => !n.hidden;

test("sidebar v2.1: panel, untracked badge, advisory warning and one card", async () => {
  const { $, settle } = await open({ ...configured, identifyOpenDoc: () => v21(), getUntracked: () => untrackedCards });
  await settle();
  assert.ok(visible($("panel")));
  assert.equal($("matter").textContent, "SPA-Nordlicht");
  assert.equal($("version").textContent, "v2.1 · counterparty");
  assert.equal($("base").textContent, "v1 (ours, 2026-11-10)");
  assert.equal($("supersedes").textContent, "v2");
  assert.equal($("untrackedBadge").textContent, "1 change without track changes");
  assert.ok(visible($("gate")) && visible($("ack")) && !visible($("acked")));
  assert.ok(visible($("openBase")));
  const cards = $("cards").querySelectorAll(".card");
  assert.equal(cards.length, 1);
  assert.match(cards[0].textContent, /Clause 4\.4 · modified/);
  assert.match(cards[0].textContent, /reject · risk high/);
});

test("sidebar: server text is shown as text, never run as HTML", async () => {
  const { $, window, settle } = await open({
    ...configured,
    identifyOpenDoc: () => { const r = v21(); r.timeline.matterId = XSS; r.flags = [XSS]; return r; },
    getUntracked: () => untrackedCards,
  });
  await settle();
  assert.equal($("matter").textContent, XSS);
  assert.equal(window.document.querySelectorAll("img").length, 0);
  assert.match($("cards").textContent, /<img src=x onerror=/);
  assert.match($("cards").textContent, /<b>red line<\/b>/);
  assert.equal(window.document.querySelectorAll("#cards b").length, 0);
  assert.equal(window.__pwned, undefined);
});

test("sidebar: acknowledge turns the warning into a dated confirmation", async () => {
  const { $, calls, settle: s } = await open({
    ...configured, identifyOpenDoc: () => v21(), getUntracked: () => untrackedCards,
    ackUntracked: () => ({ versionId: "id-v21", acknowledgedAt: "2026-11-15T10:00:00.000Z" }),
  });
  await s();
  $("ack").click();
  await s();
  assert.deepEqual(calls.find((c) => c[0] === "ackUntracked"), ["ackUntracked", "id-v21"]);
  assert.ok(!visible($("gate")) && !visible($("ack")) && visible($("acked")));
  assert.match($("acked").textContent, /^Acknowledged 2026-11-15 10:00 UTC\./);
});

test("sidebar: an acknowledged version shows no warning from the start", async () => {
  const { $, settle } = await open({
    ...configured, identifyOpenDoc: () => v21({ checked: true, count: 1, acknowledgedAt: "2026-11-14T08:30:00Z" }),
    getUntracked: () => untrackedCards,
  });
  await settle();
  assert.ok(!visible($("gate")) && visible($("acked")));
});

const clean = (after, before, extra = {}) => ({
  versionId: "id-v21", checked: true,
  units: [{ id: "U1", section: "4.4", kind: "modified", before, after, finding: { verdict: "reject", risk: "high", rationale: "x", counterProposal: null }, ...extra }],
});
const P44 = "4.4 None of the limitations in this Clause 4 apply in case of intent.";

test("sidebar: card suggests only the missing words; mark the spot, copy them", async () => {
  const { $, calls, settle: s } = await open({
    ...configured, identifyOpenDoc: () => v21(),
    getUntracked: () => clean(P44, P44.replace("intent.", "intent or fraud.")),
    placeSuggestion: () => true, goToText: () => true,
  });
  await s();
  const card = $("cards").querySelector(".card");
  assert.equal(card.querySelector(".hint").textContent, "Insert “or fraud” after “…in case of intent”");
  assert.match(card.textContent, /In Suggesting mode: Mark the spot, then paste\./);
  const buttons = [...card.querySelectorAll("button")].map((b) => b.textContent);
  assert.deepEqual(buttons, ["Go to", "Mark the spot", "Copy “or fraud”"]);

  card.querySelectorAll("button")[1].click();
  await s();
  const [, edit] = calls.find((c) => c[0] === "placeSuggestion");
  assert.deepEqual({ ...edit }, { kind: "insert", before: " 4 apply in case of intent", oldText: "", after: "." });
  assert.match(card.textContent, /Marked\..*paste \(Ctrl\+V\)\./);
  card.querySelectorAll("button")[2].click();
  assert.match(card.textContent, /Copied “or fraud”\./);
});

test("sidebar: deletion has no copy button; spot not found says so; large change keeps 'Copy our wording'", async () => {
  const del = await open({
    ...configured, identifyOpenDoc: () => v21(),
    getUntracked: () => clean("accurate on the Signing Date and on the Closing Date.", "accurate on the Signing Date."),
    placeSuggestion: () => false,
  });
  await del.settle();
  const card = del.$("cards").querySelector(".card");
  assert.deepEqual([...card.querySelectorAll("button")].map((b) => b.textContent), ["Go to", "Mark the text"]);
  assert.match(card.textContent, /Delete “and on the Closing Date”.*press Delete\./);
  card.querySelectorAll("button")[1].click();
  await del.settle();
  assert.match(card.textContent, /The spot was not found/);

  const whole = await open({
    ...configured, identifyOpenDoc: () => v21(),
    getUntracked: () => clean("Price.", `Price; ${"x ".repeat(120)}`),
  });
  await whole.settle();
  const w = whole.$("cards").querySelector(".card");
  assert.deepEqual([...w.querySelectorAll("button")].map((b) => b.textContent), ["Go to", "Copy our wording"]);
  assert.equal(w.querySelector(".hint"), null);
});

test("sidebar: go to and open base version", async () => {
  const { $, calls, settle: s } = await open({
    ...configured, identifyOpenDoc: () => v21(), getUntracked: () => untrackedCards,
    goToText: () => false,
    openBaseVersion: () => ({ url: "https://docs.google.com/document/d/gdoc-1/edit", reused: false }),
  });
  await s();
  const [go] = $("cards").querySelectorAll("button");
  go.click();
  await s();
  assert.equal(calls.find((c) => c[0] === "goToText")[1], untrackedCards.units[0].after);
  assert.match($("cards").textContent, /Not found in the document/);

  $("openBase").click();
  await s();
  assert.deepEqual(calls.find((c) => c[0] === "openBaseVersion"), ["openBaseVersion", "id-v1", "SPA-Nordlicht v1 (ours, 2026-11-10)"]);
  const a = $("baseLink").querySelector("a");
  assert.equal(a.href, "https://docs.google.com/document/d/gdoc-1/edit");
  assert.equal(a.target, "_blank");
  assert.equal(a.rel, "noopener");
});

test("sidebar: unknown document -> triage; their version cannot be registered without a base", async () => {
  let identified = 0;
  const { $, calls, window, settle: s } = await open({
    ...configured,
    identifyOpenDoc: () => (identified++ ? v21() : { status: "unknown", flags: [], timeline: null, candidates: [{ matterId: "SPA-Polaris", label: "v1", score: 0.12 }] }),
    getUntracked: () => untrackedCards,
    sendToTriage: () => ({ triageId: "T-1", candidates: [], matters: [{ id: "SPA-Nordlicht", title: `Nordlicht ${XSS}` }] }),
    getTriageBases: (_t, _m, side) => ({ bases: side === "theirs" ? [] : [{ versionId: "id-v21", label: "v2.1", coverage: 0.97 }], flags: [] }),
    assignTriage: () => ({ versionId: "id-v3" }),
  });
  await s();
  assert.ok(visible($("unknown")) && !visible($("panel")));
  assert.equal($("candidates").textContent, "SPA-Polaris: v1 (12% of paragraphs match)");
  $("triage").click();
  await s();
  assert.ok(visible($("triageForm")));
  assert.equal(window.document.querySelectorAll("img").length, 0, "matter title shown as text");

  const pick = (id, value) => { $(id).value = value; $(id).dispatchEvent(new window.Event("change")); };
  pick("tMatter", "SPA-Nordlicht");
  pick("tSide", "theirs");
  await s();
  $("tConfirm").click();
  await s();
  assert.match($("status").textContent, /Choose the base/);
  assert.ok(!calls.some((c) => c[0] === "assignTriage"), "nothing registered");

  pick("tSide", "ours");
  await s();
  assert.equal($("tBase").value, "id-v21", "best base preselected");
  $("tConfirm").click();
  await s(); await s();
  assert.deepEqual(calls.find((c) => c[0] === "assignTriage"), ["assignTriage", "T-1", "SPA-Nordlicht", "ours", "id-v21"]);
  assert.ok(visible($("panel")), "identified again after registering");
});

test("sidebar: not set up -> settings open, nothing exported; errors without the Exception prefix", async () => {
  const fresh = await open({ getSettingsInfo: () => ({ url: "", hasToken: false }) });
  assert.ok(fresh.$("settings").open);
  assert.equal(fresh.$("token").type, "password");
  assert.match(fresh.$("status").textContent, /^Not set up yet/);
  assert.ok(!fresh.calls.some((c) => c[0] === "identifyOpenDoc"));

  const failing = await open({ ...configured, identifyOpenDoc: () => { throw new Error("Exception: The server rejected the token. Check the token in Settings."); } });
  await failing.settle();
  assert.equal(failing.$("status").textContent, "The server rejected the token. Check the token in Settings.");
  assert.equal(failing.$("status").className, "error");
});

test("sidebar: saving settings clears the token field and checks the document", async () => {
  const { $, calls, settle: s } = await open({
    getSettingsInfo: () => ({ url: "", hasToken: false }),
    saveSettings: (url) => ({ url, hasToken: true }),
    identifyOpenDoc: () => v21(), getUntracked: () => untrackedCards,
  });
  $("url").value = "https://demo.trycloudflare.com";
  $("token").value = "x".repeat(40);
  $("save").click();
  await s(); await s();
  assert.deepEqual(calls.find((c) => c[0] === "saveSettings"), ["saveSettings", "https://demo.trycloudflare.com", "x".repeat(40)]);
  assert.equal($("token").value, "");
  assert.equal($("settingsMsg").textContent, "Saved.");
  assert.ok(calls.some((c) => c[0] === "identifyOpenDoc"));
});

// ---- playbook review ------------------------------------------------------------------------

const finding = (verdict, extra = {}) => ({ verdict, risk: "medium", ruleIds: ["SPA-04"], rationale: "why", counterProposal: null, comment: "Please keep our wording.", flags: [], ...extra });
const reviewUnit = (id, verdict, state, extra = {}) => ({
  id, source: "tracked", kind: "modified", section: "4. Limitation of Liability", status: "new", state,
  before: "4.1 claims below EUR 25,000.", after: "4.1 claims below EUR 100,000.", finding: finding(verdict), ...extra,
});
const findingsRes = (units) => ({ versionId: "id-v21", playbook: "spa-share-deal@2026-11", mode: "mock", model: null, reviewedAt: "2026-11-12T10:00:00Z", units });

test("review: summary, progress, document order with verdict badges; handled hidden until 'Show handled'", async () => {
  const { $, window, settle: s } = await open({
    ...configured, identifyOpenDoc: () => v21(), getUntracked: () => untrackedCards,
    getFindings: () => findingsRes([
      reviewUnit("U1", "negotiate", "open"), reviewUnit("U2", "reject", "open"),
      reviewUnit("U3", "accept", "accepted"), reviewUnit("U7", "reject", "open", { source: "untracked" }),
    ]),
  });
  await s();
  assert.ok(visible($("review")));
  assert.equal($("reviewSummary").textContent, "3 changes · 1 reject · 1 negotiate · 1 accept");
  assert.equal($("reviewProgress").textContent, "1 of 3 handled");
  assert.match($("reviewSource").textContent, /^Playbook spa-share-deal@2026-11, mock reviewer/);
  const heads = () => [...$("reviewCards").querySelectorAll(".card .verdict")].map((n) => n.textContent.split(" ")[0]);
  assert.deepEqual(heads(), ["negotiate", "reject"]);
  assert.equal($("reviewCards").querySelector(".card.c-reject .verdict b").textContent, "reject");
  assert.match($("reviewCards").textContent, /1 handled change hidden\./);
  $("showHandled").checked = true;
  $("showHandled").dispatchEvent(new window.Event("change"));
  assert.deepEqual(heads(), ["negotiate", "reject", "accept"]);
  assert.match($("reviewCards").textContent, /Accepted/);
});

test("review card: diff, steps, Go to falls back to our text, minimal edit and comment buttons", async () => {
  const { $, calls, settle: s } = await open({
    ...configured, identifyOpenDoc: () => v21(), getUntracked: () => untrackedCards,
    getFindings: () => findingsRes([reviewUnit("U1", "negotiate", "open", { finding: finding("negotiate", { counterProposal: "4.1 claims below EUR 50,000." }) })]),
    goToText: (t) => t.includes("25,000"), placeSuggestion: () => true,
  });
  await s();
  const card = $("reviewCards").querySelector(".card");
  assert.equal(card.querySelector("del").textContent, "25,000.");
  assert.equal(card.querySelector("ins").textContent, "100,000.");
  assert.deepEqual([...card.querySelectorAll("ol.steps li")].map((li) => li.textContent),
    ["Reject their suggestion (✗ in the margin).", "Then, in Suggesting mode: Replace “25,000” with “50,000”."]);
  assert.deepEqual([...card.querySelectorAll("button")].map((b) => b.textContent), ["Go to", "Mark the text", "Copy “50,000”", "Copy comment"]);

  card.querySelectorAll("button")[0].click();
  await s();
  assert.deepEqual(calls.filter((c) => c[0] === "goToText").map((c) => c[1]), ["4.1 claims below EUR 100,000.", "4.1 claims below EUR 25,000."]);
  assert.equal(card.querySelector("div.muted:last-child").textContent, "", "found on the second try");
  card.querySelectorAll("button")[1].click();
  await s();
  assert.deepEqual({ ...calls.find((c) => c[0] === "placeSuggestion")[1] }, { kind: "replace", before: "4.1 claims below EUR ", oldText: "25,000", after: "." });
  card.querySelectorAll("button")[3].click();
  assert.match(card.textContent, /Copied\. Add it as a comment on their suggestion \(Ctrl\+Alt\+M\)\./);
});

test("review: refresh asks again; server text in a card is never HTML; errors are shown in the section", async () => {
  let n = 0;
  const { $, window, calls, settle: s } = await open({
    ...configured, identifyOpenDoc: () => v21(), getUntracked: () => untrackedCards,
    getFindings: () => (n++ ? findingsRes([reviewUnit("U1", "reject", "rejected")])
      : findingsRes([reviewUnit("U1", "reject", "open", { section: XSS, finding: finding("reject", { rationale: XSS, ruleIds: [XSS] }) })])),
  });
  await s();
  assert.equal(window.document.querySelectorAll("img").length, 0);
  assert.match($("reviewCards").textContent, /<img src=x onerror=/);
  assert.equal(window.__pwned, undefined);
  $("reviewRefresh").click();
  await s();
  assert.equal(calls.filter((c) => c[0] === "getFindings").length, 2);
  assert.equal($("reviewProgress").textContent, "1 of 1 handled ✓");

  const failing = await open({ ...configured, identifyOpenDoc: () => v21(), getUntracked: () => untrackedCards,
    getFindings: () => { throw new Error("Exception: Server error: not a readable .docx"); } });
  await failing.settle();
  assert.equal(failing.$("reviewCards").textContent, "Server error: not a readable .docx");
});

test("review card: with their suggestion open, Go to lands on the unchanged part and says so", async () => {
  const { $, calls, settle: s } = await open({
    ...configured, identifyOpenDoc: () => v21(), getUntracked: () => untrackedCards,
    getFindings: () => findingsRes([reviewUnit("U1", "negotiate", "open")]),
    goToText: (t) => t === "4.1 claims below EUR", placeSuggestion: () => true,
  });
  await s();
  const card = $("reviewCards").querySelector(".card");
  card.querySelector("button").click();
  await s();
  assert.deepEqual(calls.filter((c) => c[0] === "goToText").map((c) => c[1]),
    ["4.1 claims below EUR 100,000.", "4.1 claims below EUR 25,000.", "4.1 claims below EUR"]);
  assert.match(card.textContent, /Jumped to the unchanged part of the clause\./);
});
