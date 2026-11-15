// Minimal Office.js stand-in so the task pane can run in a normal browser
// against a sample .docx (used for UI previews and smoke tests, not in Word).
/* global JSZip */
(function () {
  const log = (...a) => console.log("[word-stub]", ...a);
  let ooxml = null;
  const sample = new URLSearchParams(location.search).get("doc") || "nda_v2_counterparty_tracked.docx";

  async function loadDoc() {
    if (ooxml) return ooxml;
    const buf = await fetch(`/samples/${sample}`).then((r) => r.arrayBuffer());
    const zip = await JSZip.loadAsync(buf);
    ooxml = await zip.file("word/document.xml").async("string");
    return ooxml;
  }

  const proxy = (name) => new Proxy(function () {}, {
    get: (_t, prop) => (prop === "then" ? undefined : prop === "items" ? [] : proxy(`${name}.${String(prop)}`)),
    apply: () => { log(name + "()"); return proxy(name + "()"); },
  });

  window.Office = {
    HostType: { Word: "Word" },
    context: { requirements: { isSetSupported: () => true } },
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
        },
      };
      return fn(context);
    },
  };
})();
