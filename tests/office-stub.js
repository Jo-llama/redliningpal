// Minimal Office.js stand-in so the task pane can run in a normal browser
// against a sample .docx (used for UI previews and smoke tests, not in Word).
/* global JSZip */
(function () {
  const log = (...a) => console.log("[word-stub]", ...a);
  let ooxml = null;
  let bytes = null;
  let stamp = null;
  const sample = new URLSearchParams(location.search).get("doc") || "nda_v2_counterparty_tracked.docx";
  window.__opened = []; // documents "opened" via createDocument, for preview checks

  async function loadDoc() {
    if (ooxml) return ooxml;
    const buf = await fetch(`/samples/${sample}`).then((r) => r.arrayBuffer());
    bytes = new Uint8Array(buf);
    const zip = await JSZip.loadAsync(buf);
    ooxml = await zip.file("word/document.xml").async("string");
    const custom = zip.file("docProps/custom.xml");
    const m = custom && (await custom.async("string")).match(/name="rc-ver"[^>]*>\s*<vt:lpwstr>([^<]+)</);
    stamp = m ? m[1] : null;
    return ooxml;
  }

  const proxy = (name) => new Proxy(function () {}, {
    get: (_t, prop) => (prop === "then" ? undefined : prop === "items" ? [] : proxy(`${name}.${String(prop)}`)),
    apply: () => { log(name + "()"); return proxy(name + "()"); },
  });

  const ok = "succeeded";
  window.Office = {
    HostType: { Word: "Word" },
    AsyncResultStatus: { Succeeded: ok, Failed: "failed" },
    FileType: { Compressed: "compressed" },
    context: {
      requirements: { isSetSupported: () => true },
      document: {
        url: sample.split("/").pop(),
        // whole file as one slice
        getFileAsync: (_type, _opts, cb) => loadDoc().then(() => cb({
          status: ok,
          value: {
            sliceCount: 1,
            getSliceAsync: (_i, cb2) => cb2({ status: ok, value: { data: Array.from(bytes) } }),
            closeAsync: () => {},
          },
        })),
      },
    },
    onReady: (cb) => setTimeout(() => cb({ host: "Word" }), 0),
  };
  window.Word = {
    ChangeTrackingMode: { trackAll: "TrackAll", off: "Off" },
    run: async (fn) => {
      const xml = await loadDoc();
      const paraCount = (xml.match(/<w:p[ >]/g) || []).length;
      const items = Array.from({ length: paraCount }, (_, i) => proxy(`paragraph[${i}]`));
      const context = {
        sync: async () => {},
        document: {
          load() {}, changeTrackingMode: "Off",
          body: { getOoxml: () => ({ value: xml }), paragraphs: { load() {}, items } },
          properties: {
            customProperties: {
              getItemOrNullObject: (name) => ({ load() {}, isNullObject: name !== "rc-ver" || !stamp, value: stamp }),
            },
          },
        },
        application: {
          createDocument: (b64) => ({
            open: () => { window.__opened.push(b64.length); log(`createDocument(${b64.length} base64 chars).open()`); },
          }),
        },
      };
      return fn(context);
    },
  };
})();
