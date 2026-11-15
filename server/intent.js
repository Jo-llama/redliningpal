// What does the sender say about the attachment? Rule-based (DE/EN/PL) so it
// is deterministic and testable. The email body is untrusted input: whatever
// classifies it may only return one of these labels, never act on the text.

export const INTENTS = ["new_version", "corrected_version", "final_for_signature", "no_document_change"];

const RULES = [
  {
    intent: "corrected_version",
    patterns: [
      /disregard (our|the|my) (previous|earlier|last|draft)/i, /(attached|this) (version )?replaces/i,
      /wrong (version|file|attachment)/i, /corrected (version|draft)/i, /please use (this|the attached) (version )?instead/i,
      /korrigierte (fassung|version)/i, /bereinigte (fassung|version)/i, /ersetzt (die|unsere) (vorherige|letzte)/i,
      /bitte (die|unsere) (vorherige|letzte) (fassung|version) (ignorieren|verwerfen)/i,
      /poprawion[aey] wersj/i, /prosz[eę] zignorowa[cć]/i, /zast[eę]puje poprzedni/i,
    ],
  },
  {
    intent: "final_for_signature",
    patterns: [/execution version/i, /for signature/i, /final (agreed )?version/i, /unterschriftsfassung/i, /zur unterschrift/i, /do podpisu/i],
  },
];

export function classifyIntent({ subject, text }, hasDocument) {
  if (!hasDocument) return { intent: "no_document_change", evidence: "no Word attachment" };
  const hay = `${subject || ""}\n${(text || "").slice(0, 4000)}`;
  for (const r of RULES) {
    const hit = r.patterns.find((p) => p.test(hay));
    if (hit) return { intent: r.intent, evidence: `matched ${hit}` };
  }
  return { intent: "new_version", evidence: "default" };
}
