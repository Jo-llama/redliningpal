// Calls Claude (or the offline mock) and returns validated findings.

import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { FINDINGS_TOOL, systemPrompt, userPayload, validateFindings, mockReview } from "../public/lib/review.js";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
export const MODEL = process.env.REDLINE_MODEL || "claude-sonnet-5-5";

export function loadPlaybook(id) {
  if (!/^[\w-]+$/.test(id)) throw new Error("bad playbook id");
  return JSON.parse(readFileSync(path.join(root, "playbooks", `${id}.json`), "utf8"));
}

export function liveMode() {
  return !!process.env.ANTHROPIC_API_KEY && !process.env.REDLINE_MOCK;
}

export async function reviewUnits(playbook, units, { mock = false } = {}) {
  if (!units.length) return { mode: "none", model: null, findings: [], issues: [] };
  const live = liveMode() && !mock;
  const raw = live ? await callClaude(playbook, units) : mockReview(playbook, units);
  const { findings, issues } = validateFindings(raw, units, playbook);
  return { mode: live ? "live" : "mock", model: live ? MODEL : null, findings, issues };
}

async function callClaude(playbook, units) {
  const resp = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "content-type": "application/json",
      "x-api-key": process.env.ANTHROPIC_API_KEY,
      "anthropic-version": "2023-06-01",
    },
    body: JSON.stringify({
      model: MODEL,
      max_tokens: 8000,
      system: systemPrompt(playbook),
      tools: [FINDINGS_TOOL],
      tool_choice: { type: "tool", name: FINDINGS_TOOL.name },
      messages: [{ role: "user", content: userPayload(playbook, units) }],
    }),
  });
  if (!resp.ok) throw new Error(`Claude API ${resp.status}: ${(await resp.text()).slice(0, 300)}`);
  const data = await resp.json();
  const block = (data.content || []).find((b) => b.type === "tool_use" && b.name === FINDINGS_TOOL.name);
  if (!block) throw new Error("model returned no findings");
  return block.input;
}
