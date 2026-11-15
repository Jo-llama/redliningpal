// Feed .eml files into intake without Gmail, then print the matter timeline:
//   npm run intake -- tests/fixtures/emails/*.eml
//   npm run intake                         (just print the timeline)

import { readFileSync } from "node:fs";
import { createContext } from "../server/context.js";
import { processMessage } from "../server/intake.js";

const ctx = createContext();
for (const f of process.argv.slice(2).filter((f) => f.endsWith(".eml")).sort()) {
  const r = await processMessage(readFileSync(f), ctx);
  console.log(`${f.split("/").pop().padEnd(40)} ${r.outcome}${r.matterId ? " -> " + r.matterId : ""}${r.triageId ? " (" + r.triageId + ")" : ""}`);
}

console.log("\nTimeline");
for (const m of ctx.store.db().matters) {
  const vs = ctx.store.versions(m.id);
  if (!vs.length) continue;
  console.log(`\n${m.id}  ${m.title}`);
  for (const v of vs) {
    const base = v.baseId ? ctx.store.version(v.baseId)?.label : "-";
    const a = v.analysis;
    const info = a
      ? `${a.units.length} changes, ${a.integrity.untracked} untracked, verdicts ${JSON.stringify(a.counts)}${a.delta ? `, ${a.delta.newUnits} new since ${a.delta.supersededLabel}` : ""}`
      : "";
    console.log(`  ${v.label.padEnd(11)} ${v.side.padEnd(6)} base ${String(base).padEnd(4)} ${v.status.padEnd(10)} ${v.receivedAt.slice(0, 10)}  ${info}`);
    for (const f of v.flags || []) console.log(`              ! ${f}`);
  }
}
const open = ctx.store.openTriage();
if (open.length) {
  console.log(`\nTriage (${open.length})`);
  for (const t of open) console.log(`  ${t.id}  "${t.subject}" from ${t.from}  candidates: ${t.candidates.map((c) => `${c.matterId}:${c.score}`).join(", ") || "none"}`);
}
