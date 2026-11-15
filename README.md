# redliningPal

**Your companion for contract negotiations.** redliningPal tracks every version the counterparty sends, compares it with the right version of yours, checks each change against your negotiation playbook, and lets you approve, reject or counter it directly in Word.

The counterparty can send their changes as **tracked changes**, or as a **clean copy**. For a clean copy, the add-in compares it with your last version. Each change goes to an LLM with the playbook, and the answer passes through a deterministic harness. You then work through the findings in the task pane: **accept**, **reject**, **comment**, or **counter-propose**. Counter-proposals are written into the document as *your* tracked changes, touching only the words that differ.

![Task pane](docs/taskpane.png)

## Why it is built this way

| Concern | Design choice |
|---|---|
| The LLM may skip a change, invent a rule, or soften a red line | `validateFindings()` gives every change exactly one finding, with missing ones set to `human_review`. It drops unknown rule ids and duplicates, and keyword guardrails from the playbook can only **escalate** a verdict, never relax it. Every intervention is listed under "Harness notes". |
| Structured output | A forced tool call (`report_findings`) checked against a JSON schema. No free-text parsing. |
| API key safety | The key stays on the local server (`server.js`). The browser only calls `/api/review`. |
| Audit | The server logs one line per review: model, playbook version, verdict counts and harness interventions. Contract text is never logged. |
| Demo without a key | Mock reviewer (keyword heuristics) runs automatically when `ANTHROPIC_API_KEY` is not set. |
| Testability | All logic is in pure modules under `public/lib/`: the OOXML tracked-change parser, paragraph diff, guardrails and validator. They are tested in Node against generated sample contracts. |

## Project layout

```
manifest.xml                 Office add-in manifest (adds a button on the Review tab)
server.js                    HTTPS server: task pane, review proxy, intake API, Gmail poller
server/intake.js             Email -> matter -> version -> base -> integrity check -> review
server/routing.js            Matter assignment by stamp, thread, subject tag, sender, text
server/matching.js           Text fingerprints: base detection, renditions
server/intent.js             "corrected version" / "for signature" detection (DE/EN/PL)
server/store.js              Version registry (JSON + content-addressed files)
server/gmail-poller.js       Reads a Gmail label over IMAP
server/docx.js               Paragraph extraction and rc-ver stamp
scripts/stamp.js | intake.js Stamp an outgoing draft | feed .eml files and print the timeline
config/intake.example.json   Our addresses, matters, counterparties, routing thresholds
playbooks/spa-share-deal.json Buyer-side SPA playbook (12 rules)
samples/spa/                 SPA scenario: v1 ours, v2 markup, v2.1 "corrected", clean copy
tests/fixtures/emails/       7 email scenarios (.eml)
public/taskpane.html|css|js  Task pane UI + Office.js actions
public/lib/ooxml.js          Rebuilds before/after text from <w:ins>/<w:del>
public/lib/diff.js           Paragraph alignment (clean-copy mode), word diff, minimal edit
public/lib/review.js         Prompt, tool schema, guardrails, validator, mock reviewer
playbooks/mutual-nda.json    Sample playbook: 9 rules with preferred / fallback / never
samples/                     Sample NDA: our v1, counterparty v2 tracked, v2 clean
tests/                       Unit tests + Office.js stub for browser preview
```

## Setup, step by step

Remaining steps: **1 Install → 2 Certificates → 3 API key → 4 Start → 5 Sideload → 6 Try it**

### 1. Install
Requires Node 18+.
```bash
cd redlining-pal
npm install
npm test            # 21 tests should pass
```

### 2. Trust a localhost certificate
Office only loads add-ins over HTTPS.
```bash
npm run certs
```
Accept the OS prompt to trust the dev CA. You only need to do this once per machine.

### 3. (Optional) Add your Claude API key
Skip this step to use the mock reviewer.
```bash
export ANTHROPIC_API_KEY=sk-ant-...         # Windows PowerShell: $env:ANTHROPIC_API_KEY="sk-ant-..."
export REDLINE_MODEL=claude-sonnet-5-5      # optional, this is the default
```

### 4. Start the server
```bash
npm start
```
You should see `redliningPal on https://localhost:3000`. Open that URL plus `/taskpane.html` in your browser once to confirm the certificate is trusted.

### 5. Sideload the add-in into Word
- **Word on the web:** open a document → **Home → Add-ins → More add-ins → My add-ins → Upload My Add-in** → choose `manifest.xml`.
- **Word desktop (Windows/Mac):** run `npx office-addin-debugging start manifest.xml desktop`, or put `manifest.xml` in a shared-folder catalog.

A **redliningPal → Review redlines** button appears on the **Review** tab.

### 6. Try it with the samples
1. Open `samples/nda_v2_counterparty_tracked.docx` in Word.
2. Click **Review redlines**, keep "Tracked changes in this document", then click **Scan & review**.
3. Expected results: the residuals clause, the non-solicit and the EUR 10k liability cap come back as red-line rejects (flagged by guardrails). Term, return and governing law come back as negotiate.
4. Click **Counter-propose** on the Term finding. The counterparty's change is rejected, and your counter-proposal is inserted as your own tracked change.

To test clean-copy mode, open `samples/nda_v2_counterparty_clean.docx`, choose "A clean copy", and pick `samples/nda_v1_ours.docx` as the baseline.

### Browser preview (no Word needed)
```bash
PREVIEW=1 HTTP_ONLY=1 npm start
# open http://localhost:3000/preview.html
```
This runs the real task pane against a sample docx, with a stub in place of Office.js. It's useful for UI work and screenshots.

## Email intake and version tracking

In a negotiation the danger is not reading the redline. It is reading it against the **wrong version**. Typical problems:
- the counterparty edited an older draft
- a "corrected" version arrives the next morning
- a clean copy comes in a new thread
- somebody changes a clause with track changes switched off

Intake handles this from the mailbox.

```
email ─► attachments (.docx only; PDF/.doc/.docm flagged)
      ─► matter      stamp +100 · thread +60 · [TAG] in subject +60 · sender +20/30 · text match up to +50
                     auto if ≥ 60 and 25 ahead of the runner-up, otherwise -> triage queue
      ─► version     ours / theirs · base = OUR version they worked on (text-verified, stamp is a hint)
                     same base + we haven't replied yet -> supersedes their previous version (v2 -> v2.1)
                     identical text to a known version -> filed as a rendition, no new review
      ─► analysis    tracked changes vs base · INTEGRITY: "reject all" of their file must equal our base
                     findings reused for changes carried over from a superseded version
      ─► notify      "v2.1 (base v1, replaces v2), 13 changes, 1 UNTRACKED, 4 reject, 3 new since v2"
```

The **version graph** stores `base` and `supersedes` as two separate links. A corrected version has the same base as the version it replaces, and it is compared with that base, not with the faulty version it replaces. The delta against the replaced version only tells the lawyer what is new.

The **stamp** is a custom document property `rc-ver` holding a random UUID. It contains no names, because the counterparty can see document properties. Word keeps it when the other side edits the file. The base is still chosen by matching the text, so a stamp copied into the wrong round can't mislead the system.

### Try it without Gmail
```bash
npm run intake -- tests/fixtures/emails/*.eml
```
This prints the timeline: v1 ours, v2 superseded by v2.1 with **1 untracked change** (the fraud carve-out in 4.4 was deleted with tracking off), the clean copy filed as a rendition, and the unknown NDA in triage. Delete `data/` to start over.

### Gmail demo, step by step

Use synthetic documents only (the samples in `samples/spa/`), never real client contracts.

Remaining steps: **1 App password → 2 Label + filter → 3 Config → 4 Start → 5 Send the scenario → 6 Check**

1. **App password.** Turn on 2-step verification for your Google account, then create an app password at <https://myaccount.google.com/apppasswords>. In Gmail settings, under *Forwarding and POP/IMAP*, check that IMAP is enabled.
2. **Label and filter.** Your intake address is your own address plus a tag: `yourname+redline@gmail.com`. In Gmail, create the filter `to:(yourname+redline@gmail.com)` → *Apply label* `redline-intake`. Under *Settings → Labels*, check that "Show in IMAP" is ticked for that label.
3. **Config.** Copy `config/intake.example.json` to `config/intake.json` and change these fields:
   - `ourAddresses`: `["yourname@gmail.com"]`
   - `intakeAddresses`: `["yourname+redline@gmail.com"]`
   - in matter `SPA-Nordlicht`, `counterpartyAddresses`: your second address (it plays the Seller's counsel)
4. **Start.**
   ```bash
   export GMAIL_USER=yourname@gmail.com
   export GMAIL_APP_PASSWORD="xxxx xxxx xxxx xxxx"
   export GMAIL_POLL_SECONDS=30
   npm start
   ```
   The log should show `[gmail] polling label "redline-intake"`.
5. **Send the scenario.** Each mail goes to the other party with `yourname+redline@gmail.com` in CC:
   1. From your address to the Seller address: subject `[SPA-Nordlicht] Share Purchase Agreement`, attach `samples/spa/spa_v1_ours.docx`. It is already stamped. For your own drafts, run `npm run stamp -- draft.docx` first.
   2. From the Seller address, *Reply all*: attach `spa_v2_theirs_tracked.docx`.
   3. From the Seller address, *Reply all* again: "Please disregard our draft of yesterday, the attached version replaces it." Attach `spa_v2-1_theirs_corrected.docx`.
6. **Check.** The server log shows a `[notify]` line for each version. To see the full picture, open `https://localhost:3000/api/matters/SPA-Nordlicht`, or run `npm run intake` with no files to print the timeline.

If your own CC copy does not appear under the label, Gmail has hidden a copy addressed to yourself. Forward the sent mail to `yourname+redline@gmail.com`; intake recognises forwards of your own mails as outbound.

### API
| Method | Path | Purpose |
|---|---|---|
| POST | `/api/intake` | Raw email (RFC 822) as body. This is the hook for n8n or Power Automate later. |
| GET | `/api/matters`, `/api/matters/:id` | Matters and their version graph with analyses |
| GET | `/api/versions/:id/file` | Download the stored .docx of a version |
| GET | `/api/triage` | Mails that could not be assigned automatically |
| POST | `/api/triage/:id/assign` | `{ "matterId": "..." }`. Re-processes the mail for that matter. |

## Writing your own playbook
Add a JSON file to `playbooks/`. Each rule has these fields:
```json
{
  "id": "NDA-04", "topic": "Residuals clause",
  "preferred": "...", "fallback": null, "never": "...", "guidance": "...",
  "keywords": ["residual"],                  // mock reviewer matching
  "redFlags": ["residual", "unaided memory"], // newly added phrase => escalate
  "redFlagVerdict": "reject"                  // floor: negotiate | human_review | reject
}
```

## Requirements and known limits
- **WordApi 1.6** is needed to accept or reject tracked changes. That means Word on the web or current Microsoft 365 desktop. On older builds, clean-copy mode and comments still work.
- Changes are grouped **per paragraph**. Accept and reject apply to all tracked changes in that paragraph.
- Text boxes, footnotes, headers and footers are not scanned.
- If a counter-proposal can't be placed as a minimal edit, the add-in replaces the whole paragraph text. Character formatting inside that paragraph is then lost.
- Clean-copy alignment is heuristic (LCS on paragraphs plus word similarity). Heavily restructured documents may pair the wrong paragraphs.
- Findings are suggestions. The playbook and a lawyer decide.

Intake limits:
- The registry is a JSON file for a single user. The store interface is where Postgres, SharePoint or a DMS (iManage, NetDocuments) would plug in.
- Intent detection is rule-based (DE/EN/PL).
- A version is treated as "replacing" another only when it has the same base and we have not replied yet. If the email does not say it is a correction, the version is flagged for confirmation.

## Next steps (ideas)
- Version timeline and "open base version" button in the task pane
- "Reopened" flag: the counterparty changes a clause that was agreed in an earlier round
- LLM intent classification behind the rules (output limited to the same enum)
- n8n or Power Automate flow posting to `/api/intake`, as an alternative to the poller
- Eval set: a labelled redline corpus to score verdict accuracy per rule (see the clause-harness approach)
- Export a negotiation log (change, verdict, action taken) as .xlsx for the matter file
- Per-rule fallback ladders (first, second, final position) across negotiation rounds
- Run the model on Azure/Bedrock in an EU region for client-data compliance
