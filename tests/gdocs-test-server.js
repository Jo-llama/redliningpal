// Helper for tests/gdocs-client.test.js (not a test itself): the token-protected
// app in its own process, so the test can call it synchronously like UrlFetchApp.
// Prints the port on stdout, exits when stdin closes (the test ended or died).
import { readFileSync, mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import http from "node:http";
import path from "node:path";
import { createContext } from "../server/context.js";
import { processMessage } from "../server/intake.js";
import { createApp } from "../server/app.js";

const root = new URL("..", import.meta.url).pathname;
const ctx = {
  ...createContext({
    configFile: new URL("../config/intake.example.json", import.meta.url).pathname,
    dataDir: mkdtempSync(path.join(tmpdir(), "redline-gdocs-")),
  }),
  mock: true, notify: () => {},
};
for (const f of ["01_ours_v1.eml", "02_theirs_v2.eml", "03_theirs_v2-1_corrected.eml"]) {
  await processMessage(readFileSync(new URL(`./fixtures/emails/${f}`, import.meta.url)), ctx);
}
const server = http.createServer(createApp(ctx, { root, token: process.env.TEST_TOKEN, audit: () => {} }))
  .listen(0, "127.0.0.1", () => process.stdout.write(`${server.address().port}\n`));
process.stdin.on("end", () => process.exit(0)).resume();
