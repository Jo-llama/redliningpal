// Review logic shared by the task pane (browser) and the server (Node):
// deterministic guardrails, prompt payload, tool schema, response validation,
// and a keyword-based mock reviewer for demos without an API key.

export const VERDICTS = ["accept", "negotiate", "reject", "human_review"];
export const RISKS = ["low", "medium", "high"];

/** JSON schema the model must fill (forced tool call). */
export const FINDINGS_TOOL = {
  name: "report_findings",
  description: "Report one finding per review unit, judged strictly against the playbook.",
  input_schema: {
    type: "object",
    properties: {
      findings: {
        type: "array",
        items: {
          type: "object",
          properties: {
            unit_id: { type: "string" },
            verdict: { type: "string", enum: ["accept", "negotiate", "reject", "human_review"] },
            risk: { type: "string", enum: RISKS },
            rule_ids: { type: "array", items: { type: "string" } },
            rationale: { type: "string", description: "1-3 sentences, cite the playbook position" },
            counter_proposal: {
              type: ["string", "null"],
              description: "Full replacement text for the paragraph, or null if none",
            },
            comment: { type: "string", description: "Comment to place in the document for the counterparty" },
          },
          required: ["unit_id", "verdict", "risk", "rule_ids", "rationale", "counter_proposal", "comment"],
        },
      },
    },
    required: ["findings"],
  },
};

export function systemPrompt(playbook) {
  return [
    `You are a contract negotiation reviewer acting for ${playbook.party || "our client"}.`,
    `Contract type: ${playbook.contractType}. Governing playbook version: ${playbook.version}.`,
    "You receive the counterparty's changes as review units (before -> after).",
    "Judge every unit ONLY against the playbook rules provided. Do not invent rules.",
    "- accept: the change is within a preferred or fallback position, or is neutral drafting.",
    "- negotiate: outside the preferred position but a fallback exists; give a counter_proposal using the fallback.",
    "- reject: hits a red line or a 'never accept' position; counter_proposal restores our wording.",
    "- human_review: no rule covers it but it changes rights or obligations, or you are unsure.",
    "Cite the rule ids you relied on. If no rule applies, rule_ids is [].",
    "Comments are addressed to the counterparty: short, professional, no internal reasoning.",
    "Return exactly one finding per unit_id via the report_findings tool.",
  ].join("\n");
}

export function userPayload(playbook, units) {
  return JSON.stringify(
    {
      playbook_rules: playbook.rules.map(({ id, topic, preferred, fallback, never, guidance }) => ({
        id, topic, preferred, fallback, never, guidance,
      })),
      review_units: units.map(({ id, section, kind, before, after }) => ({ id, section, kind, before, after })),
    },
    null,
    2,
  );
}

/**
 * Deterministic red-flag scan: keywords that must surface regardless of what
 * the model says. Matches only against text the counterparty ADDED.
 */
export function guardrailHits(playbook, unit) {
  const before = (unit.before || "").toLowerCase();
  const after = (unit.after || "").toLowerCase();
  const added = addedText(unit.before, unit.after).toLowerCase();
  const hits = [];
  for (const rule of playbook.rules) {
    for (const kw of rule.redFlags || []) {
      const k = kw.toLowerCase();
      // phrase is new in this paragraph, or (for repeated phrases) appears in the added words
      const isNew = (after.includes(k) && !before.includes(k)) || added.includes(k);
      if (isNew) { hits.push({ ruleId: rule.id, keyword: kw }); break; }
    }
  }
  return hits;
}

/** Words present in `after` but not in `before` (multiset difference), joined. */
export function addedText(before, after) {
  const counts = new Map();
  for (const w of tokenize(before)) counts.set(w, (counts.get(w) || 0) + 1);
  const added = [];
  for (const w of tokenize(after)) {
    const c = counts.get(w);
    if (c) counts.set(w, c - 1); else added.push(w);
  }
  // added words keep their order, so multi-word flags still match as phrases
  return added.join(" ");
}

function tokenize(s) {
  return (s || "").toLowerCase().split(/[^\p{L}\p{N}-]+/u).filter(Boolean);
}

/**
 * Validate and repair the model output so the UI never shows something the
 * harness can't stand behind:
 *  - every unit gets exactly one finding (missing -> human_review)
 *  - unknown unit ids and duplicate findings are dropped
 *  - unknown rule ids are removed and flagged
 *  - guardrail hits can only escalate a verdict, never relax it
 */
export function validateFindings(raw, units, playbook) {
  const ruleIds = new Set(playbook.rules.map((r) => r.id));
  const byUnit = new Map();
  const issues = [];

  for (const f of Array.isArray(raw?.findings) ? raw.findings : []) {
    if (!units.some((u) => u.id === f.unit_id)) { issues.push(`dropped finding for unknown unit ${f.unit_id}`); continue; }
    if (byUnit.has(f.unit_id)) { issues.push(`dropped duplicate finding for ${f.unit_id}`); continue; }
    const verdict = VERDICTS.includes(f.verdict) ? f.verdict : "human_review";
    if (verdict !== f.verdict) issues.push(`${f.unit_id}: invalid verdict "${f.verdict}" -> human_review`);
    const cited = Array.isArray(f.rule_ids) ? f.rule_ids : [];
    const unknown = cited.filter((id) => !ruleIds.has(id));
    if (unknown.length) issues.push(`${f.unit_id}: removed unknown rule ids ${unknown.join(", ")}`);
    byUnit.set(f.unit_id, {
      unitId: f.unit_id,
      verdict,
      risk: RISKS.includes(f.risk) ? f.risk : "medium",
      ruleIds: cited.filter((id) => ruleIds.has(id)),
      rationale: String(f.rationale || ""),
      counterProposal: f.counter_proposal ? String(f.counter_proposal) : null,
      comment: String(f.comment || ""),
      flags: [],
    });
  }

  const rank = { accept: 0, negotiate: 1, human_review: 2, reject: 3 };
  return {
    issues,
    findings: units.map((u) => {
      let f = byUnit.get(u.id);
      if (!f) {
        issues.push(`${u.id}: no finding returned -> human_review`);
        f = {
          unitId: u.id, verdict: "human_review", risk: "medium", ruleIds: [],
          rationale: "The model returned no finding for this change.", counterProposal: null, comment: "", flags: [],
        };
      }
      for (const hit of guardrailHits(playbook, u)) {
        const rule = playbook.rules.find((r) => r.id === hit.ruleId);
        f.flags.push(`Red flag "${hit.keyword}" (${hit.ruleId})`);
        if (!f.ruleIds.includes(hit.ruleId)) f.ruleIds.push(hit.ruleId);
        const floor = rule.redFlagVerdict || "human_review";
        if (rank[f.verdict] < rank[floor]) {
          issues.push(`${u.id}: guardrail escalated ${f.verdict} -> ${floor} (${hit.keyword})`);
          f.verdict = floor;
          f.risk = "high";
        }
      }
      return f;
    }),
  };
}

/** Offline reviewer: keyword heuristics so the add-in demos without a key. */
export function mockReview(playbook, units) {
  return {
    findings: units.map((u) => {
      const hits = guardrailHits(playbook, u);
      const text = (u.before + " " + u.after).toLowerCase();
      const section = (u.section || "").toLowerCase();
      const scored = playbook.rules
        .map((r) => ({
          r,
          s: (r.keywords || []).reduce((a, k) => a + (text.includes(k) ? 2 : 0) + (section.includes(k) ? 1 : 0), 0),
        }))
        .filter((x) => x.s > 0)
        .sort((a, b) => b.s - a.s);
      const rule = hits.length ? playbook.rules.find((r) => r.id === hits[0].ruleId) : scored[0]?.r;
      if (!rule) {
        return {
          unit_id: u.id, verdict: "human_review", risk: "medium", rule_ids: [],
          rationale: "[mock] No playbook rule matched this change.", counter_proposal: null,
          comment: "Please explain the purpose of this change.",
        };
      }
      const hard = hits.length > 0;
      return {
        unit_id: u.id,
        verdict: hard ? "reject" : rule.fallback ? "negotiate" : "reject",
        risk: hard ? "high" : "medium",
        rule_ids: [rule.id],
        rationale: `[mock] Touches "${rule.topic}". Our position: ${rule.preferred}`,
        counter_proposal: u.before || null,
        comment: hard
          ? `We cannot accept this change. ${rule.guidance || ""}`.trim()
          : `We propose to keep our wording here. ${rule.fallback ? "Fallback: " + rule.fallback : ""}`.trim(),
      };
    }),
  };
}
