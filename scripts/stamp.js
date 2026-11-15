// Stamp our outgoing draft before sending it:
//   npm run stamp -- "SPA_v3.docx"            -> writes "SPA_v3.stamped.docx"
//   npm run stamp -- in.docx out.docx
// The stamp (rc-ver = random UUID) lets intake recognise which of our
// versions the counterparty's reply is based on. It contains no names.

import { readFileSync, writeFileSync } from "node:fs";
import crypto from "node:crypto";
import { writeStamp, readDocx } from "../server/docx.js";

const [input, output = input.replace(/\.docx$/i, ".stamped.docx")] = process.argv.slice(2);
if (!input) {
  console.error('usage: npm run stamp -- <in.docx> [out.docx]');
  process.exit(1);
}
const id = crypto.randomUUID();
const before = await readDocx(readFileSync(input));
writeFileSync(output, await writeStamp(readFileSync(input), id));
console.log(`${before.stamp ? `replaced stamp ${before.stamp} with` : "stamped"} ${id}\n-> ${output}\nSend this file (with the intake address in CC) to register it.`);
