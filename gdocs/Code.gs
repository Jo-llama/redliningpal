// redliningPal for Google Docs: Apps Script (V8) side of the Editor add-on.
//
// The sidebar never talks to the server itself; it calls the public functions
// below via google.script.run. Only this file holds the API token, so it never
// reaches the browser. Requests go to the token-protected listener that
// server.js starts on GDOCS_PORT, reached through an HTTPS tunnel.
//
// Plain functions and globals only (no modules): Apps Script loads this file as
// is, and tests/gdocs-client.test.js loads it into node:vm with fake services.
// Never log document text. Untracked-change cards carry contract text to the
// sidebar (the lawyer's own document); nothing here writes it anywhere else.
//
// Platform limits (Docs API): suggestions cannot be accepted, rejected or created
// by a script, so this add-on reads and warns; the lawyer acts in Suggesting mode.

var RP_URL_KEY = "RP_SERVER_URL";
var RP_TOKEN_KEY = "RP_API_TOKEN";
var RP_BASE_KEY = "RP_BASE_"; // + versionId -> Google Doc id of the "Open base version" copy
var RP_MIN_TOKEN = 32;
var RP_DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
var RP_GDOC = "application/vnd.google-apps.document";
var RP_ID = /^[A-Za-z0-9_-]{1,80}$/;

// ---- menu and sidebar ----------------------------------------------------------

function onOpen() {
  DocumentApp.getUi().createAddonMenu()
    .addItem("Open version panel", "showSidebar")
    .addToUi();
}

function onInstall() {
  onOpen();
}

function showSidebar() {
  var html = HtmlService.createTemplateFromFile("Sidebar").evaluate().setTitle("redliningPal");
  DocumentApp.getUi().showSidebar(html);
}

/** <?!= include("VersionView") ?> in Sidebar.html: the shared view logic. */
function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

// ---- settings (per Google user, never in the document or the code) ---------------

/** What the sidebar may know about the settings: the address, never the token. */
function getSettingsInfo() {
  var p = PropertiesService.getUserProperties();
  return { url: p.getProperty(RP_URL_KEY) || "", hasToken: !!p.getProperty(RP_TOKEN_KEY) };
}

/** url: https tunnel address; token: leave empty to keep the stored one. */
function saveSettings(url, token) {
  var clean = String(url || "").trim().replace(/\/+$/, "");
  if (!/^https:\/\/[^\s/?#]+(\/[^\s?#]*)?$/.test(clean)) {
    throw new Error("The server address must start with https:// (the tunnel address).");
  }
  var p = PropertiesService.getUserProperties();
  var t = String(token || "").trim();
  if (t) {
    if (t.length < RP_MIN_TOKEN) throw new Error("The token must be at least " + RP_MIN_TOKEN + " characters.");
    p.setProperty(RP_TOKEN_KEY, t);
  } else if (!p.getProperty(RP_TOKEN_KEY)) {
    throw new Error("Enter the API token (REDLINE_API_TOKEN of the server).");
  }
  p.setProperty(RP_URL_KEY, clean);
  return getSettingsInfo();
}

// ---- calls to the redliningPal server -----------------------------------------------

/** Stored address and token; throws before anything is exported or sent. */
function rpConfig_() {
  var p = PropertiesService.getUserProperties();
  var base = p.getProperty(RP_URL_KEY);
  var token = p.getProperty(RP_TOKEN_KEY);
  if (!base || !token) throw new Error("Not set up yet: open Settings and enter the server address and token.");
  if (!/^https:\/\//.test(base)) throw new Error("The stored server address is not https://. Save Settings again.");
  return { base: base, token: token };
}

function rpId_(id, what) {
  if (!RP_ID.test(String(id || ""))) throw new Error("Invalid " + what + ".");
  return encodeURIComponent(id);
}

/**
 * @param method  "get" | "post"
 * @param path    "/api/..."
 * @param opts    { payload?, contentType?, headers? }
 * @returns the HTTPResponse (status < 400)
 */
function rpRequest_(method, path, opts) {
  var cfg = rpConfig_();
  var headers = { Authorization: "Bearer " + cfg.token };
  var extra = (opts && opts.headers) || {};
  for (var k in extra) if (k.toLowerCase() !== "authorization") headers[k] = extra[k];
  var params = { method: method, headers: headers, muteHttpExceptions: true };
  if (opts && opts.payload !== undefined) {
    params.payload = opts.payload;
    params.contentType = opts.contentType || "application/json";
  }
  var res;
  try {
    res = UrlFetchApp.fetch(cfg.base + path, params);
  } catch (e) {
    throw new Error("Server not reachable. Are the redliningPal server and the tunnel running, and is the address in Settings current?");
  }
  var code = res.getResponseCode();
  if (code === 401) throw new Error("The server rejected the token. Check the token in Settings.");
  if (code >= 400) {
    var why = "HTTP " + code;
    try {
      var body = JSON.parse(res.getContentText());
      if (body && typeof body.error === "string") why = body.error;
    } catch (e) { /* not JSON: keep the status */ }
    throw new Error("Server error: " + why);
  }
  return res;
}

function rpApi_(method, path, opts) {
  var res = rpRequest_(method, path, opts);
  var body = null;
  try { body = JSON.parse(res.getContentText()); } catch (e) { body = null; }
  if (!body) throw new Error("The server sent no JSON. Is the address the redliningPal Google Docs port?");
  return body;
}

function rpJson_(method, path, data) {
  return rpApi_(method, path, { payload: JSON.stringify(data), contentType: "application/json" });
}

// Export needs drive.readonly: drive.file alone gets HTTP 403 for the open Doc
// (tested 2026-11-15). drive.file stays for files the add-on creates itself.

/** The open Doc as .docx, exactly like File > Download > Microsoft Word. */
function rpExportOpenDoc_() {
  var id = DocumentApp.getActiveDocument().getId();
  // The Docs export link (File > Download), not the Drive REST API: that one needs
  // the Drive API enabled in the script's hidden Cloud project (HTTP 403 otherwise).
  var url = "https://docs.google.com/document/d/" + encodeURIComponent(id) + "/export?format=docx";
  var res = UrlFetchApp.fetch(url, {
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true,
  });
  var code = res.getResponseCode();
  if (code >= 400) {
    // Google's own reason (e.g. insufficient scope, API disabled); no document content
    var reason = "";
    try { reason = String(JSON.parse(res.getContentText()).error.message || "").slice(0, 300); } catch (e) { reason = ""; }
    throw new Error("Google refused to export this document (HTTP " + code + (reason ? ": " + reason : "") + ").");
  }
  return res.getBlob().getBytes();
}

// ---- the open document --------------------------------------------------------------

/** Which matter and version is the open Doc? Same answer as for Word, plus { opened }. */
function identifyOpenDoc() {
  rpConfig_(); // not set up: do not even export the document
  var bytes = rpExportOpenDoc_();
  return rpApi_("post", "/api/documents/identify-file", { payload: bytes, contentType: "application/octet-stream" });
}

/** Changes made without track changes, as cards: { versionId, checked, units[] }. */
function getUntracked(versionId) {
  return rpApi_("get", "/api/versions/" + rpId_(versionId, "version id") + "/untracked");
}

/** The lawyer has seen the untracked changes (recorded and audited on the server). */
function ackUntracked(versionId) {
  return rpApi_("post", "/api/versions/" + rpId_(versionId, "version id") + "/ack-untracked");
}

/**
 * Select the first passage that starts like `text` (whitespace-tolerant), so the
 * Doc scrolls there. Returns false if it is not in the document (any more).
 */
function goToText(text) {
  var t = String(text || "").replace(/\s+/g, " ").trim();
  if (!t) return false;
  var words = t.slice(0, 80).split(" ");
  if (words.length > 1 && t.length > 80) words.pop(); // do not end on a cut-off word
  var pattern = words.map(function (w) { return w.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"); }).join("\\s+");
  var doc = DocumentApp.getActiveDocument();
  var found = doc.getBody().findText(pattern);
  if (!found) return false;
  var range = doc.newRange()
    .addElement(found.getElement().asText(), found.getStartOffset(), found.getEndOffsetInclusive())
    .build();
  doc.setSelection(range);
  return true;
}

// ---- open base version ------------------------------------------------------------------

/**
 * Google Doc copy of a registered version (our base), created once and reused.
 * Needs the Advanced Drive service (appsscript.json -> dependencies) to convert .docx.
 * @returns { url, reused }
 */
function openBaseVersion(versionId, name) {
  var id = rpId_(versionId, "version id");
  var props = PropertiesService.getUserProperties();
  var key = RP_BASE_KEY + versionId;
  var known = props.getProperty(key);
  if (known) {
    try {
      var f = Drive.Files.get(known, { fields: "id,trashed" });
      if (f && !f.trashed) return { url: "https://docs.google.com/document/d/" + known + "/edit", reused: true };
    } catch (e) { /* deleted or no longer accessible: create a new copy */ }
  }
  var title = String(name || "Base version").replace(/[\r\n\t]+/g, " ").slice(0, 120) + " (redliningPal copy)";
  var blob = rpRequest_("get", "/api/versions/" + id + "/file").getBlob().setContentType(RP_DOCX);
  var created;
  try {
    created = Drive.Files.create({ name: title, mimeType: RP_GDOC }, blob, { fields: "id" });
  } catch (e) {
    throw new Error("Could not create the Google Doc copy. In the Apps Script editor, add the service Drive API (Services +), then try again. (" + String(e.message || e).slice(0, 200) + ")");
  }
  props.setProperty(key, created.id);
  return { url: "https://docs.google.com/document/d/" + created.id + "/edit", reused: false };
}

// ---- unknown document: triage, confirmed by the lawyer ------------------------------------

/** File the open Doc for a decision. Nothing is registered until assignTriage(). */
function sendToTriage() {
  rpConfig_();
  var bytes = rpExportOpenDoc_();
  // header-safe: no line breaks or non-ASCII (the server sanitises again)
  var name = String(DocumentApp.getActiveDocument().getName() || "google-doc").replace(/[^\w .()-]/g, "_").slice(0, 100) + ".docx";
  return rpApi_("post", "/api/documents/triage", {
    payload: bytes, contentType: "application/octet-stream", headers: { "x-file-name": name },
  });
}

function getTriageBases(triageId, matterId, side) {
  if (side !== "ours" && side !== "theirs") throw new Error('Side must be "ours" or "theirs".');
  return rpApi_("get", "/api/triage/" + rpId_(triageId, "triage id") + "/bases?matterId=" +
    encodeURIComponent(String(matterId || "")) + "&side=" + side);
}

function assignTriage(triageId, matterId, side, baseId) {
  var body = { matterId: String(matterId || ""), side: side };
  if (baseId) body.baseId = String(baseId);
  return rpJson_("post", "/api/triage/" + rpId_(triageId, "triage id") + "/assign", body);
}
