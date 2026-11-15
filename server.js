// Local HTTPS server: task pane for the Word add-in (Office requires HTTPS,
// even on localhost), the API (see server/app.js) and the Gmail poller.
// Optional second listener for the Google Docs add-on (via a tunnel):
// GDOCS_PORT + REDLINE_API_TOKEN -> API only, bearer token required, 127.0.0.1 only.

import https from "node:https";
import http from "node:http";
import { fileURLToPath } from "node:url";
import path from "node:path";
import { liveMode, MODEL } from "./server/review-service.js";
import { createContext } from "./server/context.js";
import { createApp } from "./server/app.js";
import { startGmailPoller } from "./server/gmail-poller.js";

const root = path.dirname(fileURLToPath(import.meta.url));
const PORT = Number(process.env.PORT || 3000);
const ctx = createContext();
const app = createApp(ctx, { root });

async function start() {
  const banner = `[${liveMode() ? "live: " + MODEL : "mock reviewer - set ANTHROPIC_API_KEY for live review"}] config: ${path.relative(root, ctx.config._file)}`;
  if (process.env.HTTP_ONLY) {
    http.createServer(app).listen(PORT, () => console.log(`redliningPal on http://localhost:${PORT} ${banner}`));
  } else {
    const certs = await import("office-addin-dev-certs");
    const options = await certs.getHttpsServerOptions();
    https.createServer(options, app).listen(PORT, () => console.log(`redliningPal on https://localhost:${PORT} ${banner}`));
  }
  startGdocsListener();
  startGmailPoller(ctx);
}

function startGdocsListener() {
  const port = process.env.GDOCS_PORT;
  if (!port) return;
  // createApp refuses a missing or short token: fail loudly instead of exposing an open API
  const gdocs = createApp(ctx, { root, token: process.env.REDLINE_API_TOKEN ?? "" });
  http.createServer(gdocs).listen(Number(port), "127.0.0.1",
    () => console.log(`Google Docs API (token required) on http://127.0.0.1:${port} - point the tunnel here`));
}

start();
