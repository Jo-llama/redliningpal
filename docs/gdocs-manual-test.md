# Google Docs add-on: manual test checklist

What the automated tests cannot reach: a real Google account, Google's export and
conversion, OAuth consent, the tunnel. Run it after changes to `gdocs/`, the
token listener or the export path. Synthetic samples only (`samples/spa/`).

Result column: ✅ pass · ❌ fail (note what happened) · – not run.
Last run: 2026-11-15, Chrome on Windows, personal Google account, cloudflared quick tunnel.

## Setup

```bash
export REDLINE_DATA=~/redline-gdocs-demo            # keeps data/ clean
npm run intake -- tests/fixtures/emails/*.eml       # SPA-Nordlicht v1, v2, v2.1, clean copy
export REDLINE_API_TOKEN=$(node -e 'console.log(require("crypto").randomBytes(24).toString("base64url"))')
HTTP_ONLY=1 GDOCS_PORT=3001 npm start
# second terminal
~/.local/bin/cloudflared tunnel --url http://127.0.0.1:3001
```

Add-on: standalone project at script.google.com with the four files from `gdocs/`
(`npm run gdocs:push`, or paste them), then *Deploy → Test deployments → Editor
Add-on*, one test per Doc, config *Installed and enabled*, *Latest code*.

## A. Google Docs platform behaviour (feasibility, run once)

| # | Check | Expected | Result |
|---|---|---|---|
| A1 | Upload `spa_v2_theirs_tracked.docx`, File → Save as Google Docs | Tracked changes become suggestions, author "Northwind Legal", "From imported document"; a replacement is one suggestion | ✅ |
| A2 | File → Details of the converted Doc | No `rc-ver` property visible | ✅ (not visible) |
| A3 | File → Download → .docx, compare with the original (`readDocx`) | Same 11 ins / 11 del, all 27 paragraphs equal before and after, `rc-ver` stamp still present | ✅ |
| A4 | Clause numbers ("4.4") | Plain text, not list numbering | ✅ |

## B. Tunnel and token

| # | Check | Expected | Result |
|---|---|---|---|
| B1 | `curl <tunnel>/api/matters` | 401 | ✅ |
| B2 | Same with `Authorization: Bearer $REDLINE_API_TOKEN` | 200 | ✅ |
| B3 | `curl -H "Authorization: Bearer …" <tunnel>/taskpane.html` | 404 (API only) | ✅ |
| B4 | `GDOCS_PORT` set, no token | Server refuses to start | ✅ |
| B5 | Word port 3000 without token | 200, unchanged | ✅ |

## C. Add-on: first run

| # | Check | Expected | Result |
|---|---|---|---|
| C1 | Open version panel before Settings | "Not set up yet…", Settings open, nothing exported | ✅ |
| C2 | Consent screen | Unverified-app warning (own script); scopes: current document, Drive read-only, Drive files of this app, external requests, sidebar. Nothing else (no Gmail, no full Drive write) | ✅ |
| C3 | Save Settings (https tunnel address + token) | "Saved.", token field empty, "(stored; leave empty to keep)" | ✅ |
| C4 | Address `http://…` | Refused: must start with https:// | – |
| C5 | Wrong token | "The server rejected the token. Check the token in Settings." | – |
| C6 | Stop the tunnel, Check again | "Server not reachable. Are the redliningPal server and the tunnel running…" | – |

## D. Version panel

| # | Doc | Expected | Result |
|---|---|---|---|
| D1 | Converted v2 (or its .docx round trip) | SPA-Nordlicht · v2 · counterparty · matches the registered file · base v1 (ours, 2026-11-10) · supersedes none · warning "Superseded by v2.1 … outdated version" | ✅ |
| D2 | Base copy of v1 (created in E1) | v1 · our version · base none (first draft) · "A newer version exists: v2.1 (theirs…)" | ✅ |
| D3 | Converted v2.1 | v2.1 · counterparty · base v1 · supersedes v2 · no outdated warning | ✅ |
| D4 | Server audit log for D1–D3 | `identify_file` lines with ids, score, paragraph count, `stamped`; no contract text | ✅ |

## E. Open base version

| # | Check | Expected | Result |
|---|---|---|---|
| E1 | Open base version on v2 | Link "Open v1 (ours, 2026-11-10) in a new tab"; Google Doc "SPA-Nordlicht v1 (ours, 2026-11-10) (redliningPal copy)", formatted, no suggestions; 4.1 EUR 25,000, 4.3 30% | ✅ link and copy created; text = v1 (D2 exact); formatting not checked |
| E2 | Click again | Same copy, no second file in Drive | – |
| E3 | Trash the copy, click again | New copy created | – |

## F. Untracked changes (v2.1)

| # | Check | Expected | Result |
|---|---|---|---|
| F1 | Panel | Red badge "1 change without track changes", red "Do not use Accept all yet" | ✅ |
| F2 | Card | Clause 4 · modified · REJECT · RISK HIGH; their text "…in case of intent."; our base "…in case of intent or fraud." | ✅ |
| F3 | Go to | Doc scrolls to 4.4, passage selected | ✅ |
| F4 | Card shows "Suggested edit: Insert “or fraud” after “…in case of intent”" | Buttons Go to · Mark the spot · Copy “or fraud” | ✅ |
| F4a | Suggesting mode on, Mark the spot, Copy, paste (Ctrl+V) | A suggestion that inserts only "or fraud" before the full stop, under the lawyer's name. Check whether the cursor survives the switch from the sidebar to the document (if not: note it) | ✅ suggestion inserted only "or fraud" (whether a click into the Doc was needed was not recorded) |
| F4b | (before the minimal edit, 2026-11-15) Copy our wording, paste over 4.4 | Suggestion replacing the whole sentence | ✅ |
| F5 | I have reviewed these changes | Green "Acknowledged <date> UTC"; server audit `ack_untracked` | ✅ |
| F6 | Check again after F5 | Stays green (acknowledgement is on the version, not the session) | – |

A large change (over 200 characters) still shows "Copy our wording" for the whole text.

## G. Triage of an unknown Doc

| # | Check | Expected | Result |
|---|---|---|---|
| G1 | Converted `samples/nda_v2_counterparty_tracked.docx` | "This document is not registered…", Send to triage | – |
| G2 | Send to triage | Form: matter, side, base. Nothing registered yet (`GET /api/matters` unchanged) | – |
| G3 | Side "Counterparty's", no base chosen, Register | "Choose the base…", nothing registered | – |
| G4 | Choose matter, side, base, Register | Panel shows the new version and its base; audit `triage_assign` | – |

## H. Limits (expected, documented)

- Accept all / reject in Google Docs cannot be blocked or performed by the add-on (no API for suggestions); the warning is advisory.
- Counter-proposals are pasted by the lawyer in Suggesting mode.
- The tunnel address changes on every cloudflared restart: update Settings.
