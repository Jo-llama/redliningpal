// Polls a Gmail label over IMAP and feeds each unread message into intake.
// Gmail exposes labels as IMAP folders. Needs an app password (Google account
// with 2-step verification): https://myaccount.google.com/apppasswords

import { ImapFlow } from "imapflow";
import { processMessage } from "./intake.js";

export function startGmailPoller(ctx, {
  user = process.env.GMAIL_USER,
  pass = process.env.GMAIL_APP_PASSWORD,
  label = process.env.GMAIL_LABEL || "redline-intake",
  everyMs = Number(process.env.GMAIL_POLL_SECONDS || 60) * 1000,
} = {}) {
  if (!user || !pass) {
    console.log("[gmail] GMAIL_USER / GMAIL_APP_PASSWORD not set - poller off (use POST /api/intake or npm run intake)");
    return null;
  }
  let running = false;

  async function pollOnce() {
    if (running) return;
    running = true;
    const client = new ImapFlow({ host: "imap.gmail.com", port: 993, secure: true, auth: { user, pass }, logger: false });
    try {
      await client.connect();
      const lock = await client.getMailboxLock(label);
      const batch = [];
      try {
        // collect first: no other IMAP commands are allowed while a fetch is streaming
        for await (const m of client.fetch({ seen: false }, { uid: true, source: true })) batch.push({ uid: m.uid, source: m.source });
        for (const m of batch) {
          try {
            const report = await processMessage(m.source, ctx);
            console.log(`[gmail] uid ${m.uid}: ${report.outcome}${report.matterId ? " -> " + report.matterId : ""}`);
            await client.messageFlagsAdd(String(m.uid), ["\\Seen"], { uid: true });
          } catch (err) {
            // leave unread so the next poll retries; log without contract content
            console.error(`[gmail] uid ${m.uid} failed: ${err.message}`);
          }
        }
      } finally {
        lock.release();
      }
      await client.logout();
    } catch (err) {
      console.error(`[gmail] poll failed: ${err.message}`);
      try { await client.logout(); } catch { /* ignore */ }
    } finally {
      running = false;
    }
  }

  console.log(`[gmail] polling label "${label}" for ${user} every ${everyMs / 1000}s`);
  pollOnce();
  const timer = setInterval(pollOnce, everyMs);
  return { stop: () => clearInterval(timer), pollOnce };
}
