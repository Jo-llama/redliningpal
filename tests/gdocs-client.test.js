// gdocs/Code.gs in node:vm with fake Apps Script services (UrlFetchApp,
// PropertiesService, DocumentApp, ScriptApp). The last test runs it against the
// real token-protected app (own process, tests/gdocs-test-server.js), with
// UrlFetchApp made synchronous via a child process per request.
import { test, before, after } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { spawn, spawnSync } from "node:child_process";
import vm from "node:vm";
import JSZip from "jszip";
import { buildVersionView, OUT as VIEW_OUT } from "../scripts/build-gdocs.js";
import { versionView } from "../public/lib/version-view.js";
const CODE = readFileSync(new URL("../gdocs/Code.gs", import.meta.url), "utf8");
const TOKEN = "gdocs-test-token-".padEnd(40, "x");
const EXPORT = readFileSync(new URL("./fixtures/gdocs/spa_v2_gdocs_export.docx", import.meta.url));

const response = (code, body, bytes) => ({
  getResponseCode: () => code,
  getContentText: () => (body === null ? Buffer.from(bytes || []).toString("utf8") : typeof body === "string" ? body : JSON.stringify(body)),
  getBlob: () => {
    const blob = { bytes: Array.from(bytes || []), type: null, getBytes: () => blob.bytes };
    blob.setContentType = (t) => { blob.type = t; return blob; };
    return blob;
  },
});

/** Fake of the open Doc: findText over `text`, selection recorded. */
function fakeDoc(text = "") {
  const doc = {
    selection: null, patterns: [],
    getId: () => "doc-123",
    getName: () => "Escrow draft\n(copy)",
    getBody: () => ({
      findText: (pattern) => {
        doc.patterns.push(pattern);
        const m = new RegExp(pattern).exec(text);
        return m && { getElement: () => ({ asText: () => "text-el" }), getStartOffset: () => m.index, getEndOffsetInclusive: () => m.index + m[0].length - 1 };
      },
    }),
    newRange: () => {
      const r = { parts: [], addElement: (e, a, b) => { r.parts.push([e, a, b]); return r; }, build: () => r };
      return r;
    },
    setSelection: (r) => { doc.selection = r.parts; },
  };
  return doc;
}

/** Fake Advanced Drive service: files by id, `trashed` per file. */
function fakeDrive() {
  const drive = {
    files: {}, created: [],
    Files: {
      get: (id) => { if (!drive.files[id]) throw new Error("File not found: " + id); return { id, trashed: drive.files[id].trashed }; },
      create: (meta, blob) => {
        const id = `gdoc-${drive.created.length + 1}`;
        drive.files[id] = { trashed: false };
        drive.created.push({ id, meta, type: blob.type, bytes: blob.getBytes() });
        return { id };
      },
    },
  };
  return drive;
}

/**
 * @param fetchImpl (url, params) -> response; calls are recorded in `calls`
 * @param props     initial user properties
 */
function load(fetchImpl, props = {}, { doc = fakeDoc(), drive = fakeDrive() } = {}) {
  const store = { ...props };
  const calls = [];
  const sandbox = {
    UrlFetchApp: { fetch: (url, params) => { calls.push({ url, params }); return fetchImpl(url, params); } },
    PropertiesService: {
      getUserProperties: () => ({
        getProperty: (k) => (k in store ? store[k] : null),
        setProperty: (k, v) => { store[k] = v; },
      }),
    },
    DocumentApp: { getActiveDocument: () => doc },
    Drive: drive,
    ScriptApp: { getOAuthToken: () => "google-oauth-token" },
  };
  vm.createContext(sandbox);
  vm.runInContext(CODE, sandbox, { filename: "Code.gs" });
  return { gs: sandbox, calls, store, doc, drive };
}
const configured = { RP_SERVER_URL: "https://demo.trycloudflare.com", RP_API_TOKEN: TOKEN };
const exportOk = (url, params) => (url.startsWith("https://docs.google.com/") ? response(200, "", EXPORT) : null);

// ---- settings ------------------------------------------------------------------

test("settings: only https addresses and long tokens are stored; the token never comes back", () => {
  const { gs, store } = load(() => assert.fail("no request expected"));
  assert.throws(() => gs.saveSettings("http://demo.trycloudflare.com", TOKEN), /https:\/\//);
  assert.throws(() => gs.saveSettings("javascript:alert(1)", TOKEN), /https:\/\//);
  assert.throws(() => gs.saveSettings("https://demo.trycloudflare.com", "short"), /at least 32/);
  assert.throws(() => gs.saveSettings("https://demo.trycloudflare.com", ""), /Enter the API token/);
  assert.deepEqual({ ...store }, {}, "nothing stored after failures");

  const info = gs.saveSettings(" https://demo.trycloudflare.com/ ", TOKEN);
  assert.deepEqual({ ...info }, { url: "https://demo.trycloudflare.com", hasToken: true });
  assert.ok(!JSON.stringify(gs.getSettingsInfo()).includes(TOKEN));
  gs.saveSettings("https://other.trycloudflare.com", ""); // empty token keeps the stored one
  assert.equal(store.RP_API_TOKEN, TOKEN);
});

// ---- identify -------------------------------------------------------------------

test("identify: refuses before setup, sends nothing", () => {
  const { gs, calls } = load(() => assert.fail("no request expected"));
  assert.throws(() => gs.identifyOpenDoc(), /Not set up yet/);
  assert.equal(calls.length, 0);
});

test("identify: exports the Doc with Google's token, posts the bytes with ours", () => {
  const { gs, calls } = load((url, p) => exportOk(url, p) || response(200, { status: "exact", versionId: "v-2" }), configured);
  assert.equal(gs.identifyOpenDoc().versionId, "v-2");
  const [exp, post] = calls;
  assert.equal(exp.url, "https://docs.google.com/document/d/doc-123/export?format=docx");
  assert.equal(exp.params.headers.Authorization, "Bearer google-oauth-token");
  assert.equal(post.url, "https://demo.trycloudflare.com/api/documents/identify-file");
  assert.equal(post.params.method, "post");
  assert.equal(post.params.headers.Authorization, `Bearer ${TOKEN}`);
  assert.equal(post.params.contentType, "application/octet-stream");
  assert.equal(post.params.payload.length, EXPORT.length);
  assert.ok(!post.params.headers.Authorization.includes("google-oauth-token"), "Google token never goes to our server");
});

test("identify: clear messages for 401, server errors, unreachable tunnel, export refused", () => {
  const msg = (impl) => { try { load(impl, configured).gs.identifyOpenDoc(); } catch (e) { return e.message; } return "no error"; };
  assert.match(msg((u, p) => exportOk(u, p) || response(401, { error: "missing or invalid API token" })), /rejected the token/);
  assert.match(msg((u, p) => exportOk(u, p) || response(400, { error: "not a readable .docx" })), /^Server error: not a readable \.docx$/);
  assert.match(msg((u, p) => exportOk(u, p) || response(502, "<html>Bad gateway</html>")), /^Server error: HTTP 502$/);
  assert.match(msg((u, p) => exportOk(u, p) || response(200, "<html>tunnel page</html>")), /no JSON/);
  assert.match(msg((u, p) => { if (exportOk(u, p)) return exportOk(u, p); throw new Error("DNS error: demo.trycloudflare.com"); }), /not reachable/);
  assert.match(msg(() => response(404, "")), /^Google refused to export this document \(HTTP 404\)\.$/);
  assert.match(msg(() => response(403, { error: { code: 403, message: "Request had insufficient authentication scopes." } })),
    /\(HTTP 403: Request had insufficient authentication scopes\.\)/);
});

test("identify: a stored http:// address (edited by hand) is refused, nothing sent to it", () => {
  const { gs, calls } = load(exportOk, { ...configured, RP_SERVER_URL: "http://demo.trycloudflare.com" });
  assert.throws(() => gs.identifyOpenDoc(), /not https/);
  assert.ok(calls.every((c) => c.url.startsWith("https://docs.google.com/")));
});

test("sidebar: server data is rendered as text, the token stays out of the page", () => {
  const html = readFileSync(new URL("../gdocs/Sidebar.html", import.meta.url), "utf8");
  assert.ok(!/innerHTML|insertAdjacentHTML|document\.write/.test(html));
  assert.ok(!/RP_API_TOKEN|getProperty/.test(html));
});

// ---- end to end: Code.gs -> real token app -----------------------------------------

let child, baseUrl;
before(async () => {
  child = spawn(process.execPath, [new URL("./gdocs-test-server.js", import.meta.url).pathname], {
    env: { ...process.env, TEST_TOKEN: TOKEN }, stdio: ["pipe", "pipe", "inherit"],
  });
  const port = await new Promise((resolve, reject) => {
    child.stdout.once("data", (d) => resolve(String(d).trim()));
    child.once("exit", (c) => reject(new Error(`test server exited (${c})`)));
  });
  baseUrl = `http://127.0.0.1:${port}`;
});
after(() => { child.stdin.end(); child.kill(); });

/** UrlFetchApp is synchronous: do the real request in a child process. */
const CHILD = `
let s = ""; process.stdin.on("data", (d) => (s += d)).on("end", async () => {
  const q = JSON.parse(s);
  const r = await fetch(q.url, { method: q.method, headers: q.headers, body: q.body ? Buffer.from(q.body, "base64") : undefined });
  process.stdout.write(JSON.stringify({ code: r.status, body: Buffer.from(await r.arrayBuffer()).toString("base64") }));
});`;
function syncFetch(url, p) {
  const headers = { ...p.headers, ...(p.contentType ? { "content-type": p.contentType } : {}) };
  const raw = p.payload === undefined ? null : typeof p.payload === "string" ? Buffer.from(p.payload, "utf8") : Buffer.from(Uint8Array.from(p.payload));
  const body = raw && raw.toString("base64");
  const out = spawnSync(process.execPath, ["-e", CHILD], { input: JSON.stringify({ url, method: p.method.toUpperCase(), headers, body }), timeout: 10000 });
  if (out.status !== 0) throw new Error("fetch failed"); // like UrlFetchApp on a network error
  const { code, body: b64 } = JSON.parse(out.stdout.toString());
  return response(code, null, Buffer.from(b64, "base64"));
}

test("end to end: Code.gs identifies the Google Docs export as v2 via the token app", () => {
  // the stored address must be https; the test server is http, so rewrite at the fetch boundary
  const fake = "https://tunnel.test";
  const { gs } = load((url, p) => exportOk(url, p) || syncFetch(url.replace(fake, baseUrl), p), { ...configured, RP_SERVER_URL: fake });
  const r = structuredClone(gs.identifyOpenDoc()); // vm-realm objects -> this realm for deepEqual
  assert.equal(r.status, "exact");
  assert.equal(r.timeline.current.label, "v2");
  assert.equal(r.timeline.base.label, "v1");
  assert.deepEqual(r.timeline.supersededBy.map((v) => v.label), ["v2.1"]);
  assert.equal(r.opened.hasTrackedChanges, true);

  const wrong = load((url, p) => exportOk(url, p) || syncFetch(url.replace(fake, baseUrl), p), { RP_SERVER_URL: fake, RP_API_TOKEN: "w".repeat(40) });
  assert.throws(() => wrong.gs.identifyOpenDoc(), /rejected the token/);
});

// ---- step 4: shared view, go to, open base version, triage ------------------------------

const sample = (p) => readFileSync(new URL(`../samples/${p}`, import.meta.url));
const exportOf = (buf) => (url) => (url.startsWith("https://docs.google.com/") ? response(200, "", buf) : null);

test("build: gdocs/VersionView.html is generated from version-view.js and up to date", () => {
  const built = buildVersionView();
  assert.equal(readFileSync(VIEW_OUT, "utf8"), built, "stale: run npm run gdocs:build");
  assert.ok(!/^\s*(import|export)\s/m.test(built), "no module syntax for Apps Script");
  const script = built.match(/<script>([\s\S]*)<\/script>/)[1];
  const box = {};
  vm.createContext(box);
  vm.runInContext(script + "\nthis.RPView = RPView;", box);
  const identify = {
    status: "probable", versionId: "a", score: 0.93, flags: [],
    timeline: {
      matterId: "M", current: { id: "a", label: "v2", side: "theirs", receivedAt: "2026-11-12T09:00:00Z", hasTrackedChanges: true },
      openedRendition: false, base: { id: "b", label: "v1", side: "ours", receivedAt: "2026-11-10T09:00:00Z" },
      supersedes: [], supersededBy: [], newerExists: false, latest: null,
      untracked: { checked: true, count: 2, acknowledgedAt: null }, chain: [], flags: [],
    },
  };
  assert.deepEqual(structuredClone(box.RPView.versionView(identify, { hasTrackedChanges: false })), versionView(identify, { hasTrackedChanges: false }));
});

test("go to: escapes regex characters, tolerates whitespace, cuts long text at a word", () => {
  const text = "4.1 The Seller is not liable for claims below EUR  100,000 (de minimis); see [Annex 3.1] $ terms.";
  const { gs, doc } = load(() => assert.fail("no request expected"), {}, { doc: fakeDoc(text) });
  assert.equal(gs.goToText("4.1 The Seller is not liable for claims below EUR 100,000 (de minimis); see [Annex 3.1] $ terms."), true);
  assert.ok(!doc.patterns[0].includes("terms"), "cut at 80 characters, on a word boundary");
  assert.match(doc.patterns[0], /\\\(de\\s\+minimis\\\);/);
  assert.equal(doc.selection[0][1], 0);
  assert.equal(gs.goToText("5.1 Claims become time-barred"), false);
  assert.equal(gs.goToText("   "), false);
  assert.equal(doc.patterns.length, 2, "empty text: no search");
});

test("open base version: converts to a Google Doc once, reuses it, recreates a deleted copy", () => {
  const v1 = sample("spa/spa_v1_ours.docx");
  const t = load((url) => response(200, null, v1), configured);
  const first = t.gs.openBaseVersion("ver-1", "SPA-Nordlicht v1 (ours, 2026-11-10)\n");
  assert.deepEqual({ ...first }, { url: "https://docs.google.com/document/d/gdoc-1/edit", reused: false });
  const [c] = t.drive.created;
  assert.equal(c.meta.mimeType, "application/vnd.google-apps.document");
  assert.equal(c.meta.name, "SPA-Nordlicht v1 (ours, 2026-11-10)  (redliningPal copy)");
  assert.equal(c.type, "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
  assert.equal(c.bytes.length, v1.length);
  assert.equal(t.calls[0].url, "https://demo.trycloudflare.com/api/versions/ver-1/file");
  assert.equal(t.calls[0].params.headers.Authorization, `Bearer ${TOKEN}`);

  assert.equal(t.gs.openBaseVersion("ver-1", "x").reused, true);
  assert.equal(t.calls.length, 1, "reused: nothing downloaded");
  t.drive.files["gdoc-1"].trashed = true;
  assert.equal(t.gs.openBaseVersion("ver-1", "x").url, "https://docs.google.com/document/d/gdoc-2/edit");
  delete t.drive.files["gdoc-2"];
  assert.equal(t.gs.openBaseVersion("ver-1", "x").reused, false);
  assert.equal(t.drive.created.length, 3);

  assert.throws(() => t.gs.openBaseVersion("../matters", "x"), /Invalid version id/);
  const broken = load(() => response(200, null, v1), configured, { drive: { Files: { create: () => { throw new Error("Drive API disabled"); } } } });
  assert.throws(() => broken.gs.openBaseVersion("ver-1", "x"), /add the service Drive API.*Drive API disabled/);
});

test("triage: header-safe file name, side and ids validated before anything is sent", () => {
  const t = load((url, p) => exportOk(url, p) || response(201, { triageId: "T-1", candidates: [], matters: [] }), configured);
  t.gs.sendToTriage();
  assert.equal(t.calls[1].params.headers["x-file-name"], "Escrow draft_(copy).docx");
  assert.equal(t.calls[1].params.headers.Authorization, `Bearer ${TOKEN}`);
  assert.throws(() => t.gs.getTriageBases("T-1", "SPA-Nordlicht", "mine"), /Side must be/);
  assert.throws(() => t.gs.assignTriage("T-1/../x", "SPA-Nordlicht", "ours"), /Invalid triage id/);
  assert.equal(t.calls.length, 2);
});

test("end to end: v2.1 -> untracked card, acknowledge opens the gate; base file reaches Drive", () => {
  const fake = "https://tunnel.test";
  const v21 = sample("spa/spa_v2-1_theirs_corrected.docx");
  const { gs, drive } = load((url, p) => exportOf(v21)(url) || syncFetch(url.replace(fake, baseUrl), p), { ...configured, RP_SERVER_URL: fake });
  const r = structuredClone(gs.identifyOpenDoc());
  const v = versionView(r, r.opened);
  assert.equal(v.title, "v2.1 · counterparty");
  assert.equal(v.gateClosed, true);
  assert.equal(v.untracked.text, "1 change without track changes");

  const u = structuredClone(gs.getUntracked(v.versionId));
  assert.equal(u.units.length, 1);
  assert.ok(u.units[0].after && u.units[0].before !== u.units[0].after);
  assert.ok(gs.ackUntracked(v.versionId).acknowledgedAt);
  const again = structuredClone(gs.identifyOpenDoc());
  assert.equal(versionView(again, again.opened).gateClosed, false);

  gs.openBaseVersion(v.base.id, `${v.matterId} ${v.base.text}`);
  const bytes = Buffer.from(drive.created[0].bytes);
  assert.equal(bytes.subarray(0, 2).toString(), "PK", "a .docx (zip) went to Drive");
  assert.ok(bytes.equals(sample("spa/spa_v1_ours.docx")), "the registered v1 file");
});

test("end to end: a counter-draft goes through triage and becomes v3 on base v2.1", async () => {
  const zip = await JSZip.loadAsync(sample("spa/spa_v2-1_theirs_corrected.docx"));
  const xml = await zip.file("word/document.xml").async("string");
  zip.file("word/document.xml", xml.replace(/<w:sectPr[\s>]/, (m) => "<w:p><w:r><w:t>15. Escrow: 10 percent held for 18 months.</w:t></w:r></w:p>" + m));
  const draft = await zip.generateAsync({ type: "nodebuffer" });
  const fake = "https://tunnel.test";
  const { gs } = load((url, p) => exportOf(draft)(url) || syncFetch(url.replace(fake, baseUrl), p), { ...configured, RP_SERVER_URL: fake });

  const t = structuredClone(gs.sendToTriage());
  assert.match(t.triageId, /^T-/);
  assert.ok(t.matters.some((m) => m.id === "SPA-Nordlicht"));
  const bases = structuredClone(gs.getTriageBases(t.triageId, "SPA-Nordlicht", "ours"));
  assert.equal(bases.bases[0].label, "v2.1");
  const done = structuredClone(gs.assignTriage(t.triageId, "SPA-Nordlicht", "ours", bases.bases[0].versionId));
  assert.equal(done.timeline.current.label, "v3");
  assert.equal(done.timeline.base.label, "v2.1");
  const r = structuredClone(gs.identifyOpenDoc());
  assert.equal(r.status, "exact");
  assert.equal(r.timeline.current.label, "v3");
});
