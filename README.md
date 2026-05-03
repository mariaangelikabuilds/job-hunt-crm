# Job Hunt CRM

A Google Workspace CRM for tracking my own job hunt. Built in Apps Script with the Anthropic API doing fit-scoring, JD extraction, email parsing, resume tailoring, cover letter drafting, and interview prep. Lives entirely inside Sheets, Drive, Gmail, and Calendar, with a mobile PWA on top.

I built this to track my own pivot from cybersecurity awareness UX (Cywareness) to Product Design Engineer roles. It runs against my actual applications. The numbers in this README come from real use.

## Try it yourself

> **Make a copy:** [open the template](https://docs.google.com/spreadsheets/d/REPLACE_WITH_TEMPLATE_SHEET_ID/copy)

Clicking that creates a copy of the Sheet, the bound Apps Script project, all `.gs` files, and the HTML files in your own Google account. Then:

1. Open the copied Sheet. You may need to wait 5-10 seconds for the menu to register.
2. The first time you click any item under the **Job CRM** menu, Google asks you to authorize the project. Review the OAuth scopes (Sheets, Gmail, Drive, Documents, Calendar, External Request) and approve.
3. Run **Job CRM → Onboarding wizard**. Provide your Anthropic API key, your master resume Doc, and your Drive parent folder. The other prompts (Slack, OpenAI) are optional.
4. Click **Open sidebar** to start.

Your data, your API key, your sheet. Nothing is shared back to me.

## What it does

**Scoring and angle generation**
- Score JDs against a master resume. Returns a 0-100 fit score, three opinionated cover-letter angles, red flags, and a one-paragraph rationale. Backed by `claude-opus-4-7` with adaptive thinking.

**Artifact generation (per application)**
- Tailor the master resume for a specific JD. Anti-fabrication rules in the system prompt: no skills not in master, no metrics not in master, no roles not in master. Honest gaps surfaced as `GAP:` notes. Output written to a Google Doc in the company's folder.
- Draft a cover letter using the picked angle as the spine. Three short paragraphs. Banned openers (no "I am writing to apply", "Excited to apply"), banned closers (no "Thank you for your consideration"), banned vocabulary list. Must reference one verifiable thing about the company; can't invent specifics.

**Interview prep mode (sidebar switches automatically when Status = Interview Scheduled)**
- Company research brief: snapshot, recent signals, leadership notes, public values, vibe signal. Uncertain claims marked `(unverified)`.
- Five to seven likely questions with response angles grounded in the candidate's actual experience.
- Persistent prep checklist (resume locked, portfolio tested, questions drafted, route checked, etc.).
- Notes textarea synced to the sheet.

**Inbox automation**
- Auto-import application emails. Label any inbox thread `job-apply` and the script appends a row, pulling Company, Role, JD link, and status from the message via Claude.
- Daily follow-up nudge. Once a day, a digest email lists applications stuck on `Applied` for more than seven days. Optional Slack delivery (incoming webhook) instead of or in addition to email.

**Slack integration (optional)**
- Daily nudge can post to a Slack channel via incoming webhook. Block Kit digest with a row link per overdue application.
- Slash command `/jobcrm log Company | Role | URL` logs a new application from any Slack channel. Returns ephemeral confirmation with a row link.
- Verification via Slack token + workspace `team_id` allowlist. (Apps Script `doPost` cannot read HTTP headers, so HMAC signature verification is not available; documented limitation below.)

**Similar past applications (optional)**
- "Find similar past applications" surfaces the top three previously-scored rows whose JDs are closest to the current row.
- Embeddings via OpenAI `text-embedding-3-small`, stored as base64-packed Float32 in a hidden column. Cosine similarity reduces to a dot product (vectors are unit-normalized).
- Useful for: spotting duplicate applications, reusing tailored-resume framings across similar roles, pattern-matching red flags across companies.

**Workspace integration**
- Per-company Drive folders. Resume variants, cover letters, screenshots all land in one place.
- Calendar integration. Flip a row's Status to `Interview Scheduled` and put a date in Next Action; a Calendar event appears.
- JD scraping. Paste a JD link with no body; the scraper fetches the page and extracts the JD via Claude. Works on Greenhouse, Lever, Workable, Ashby. LinkedIn and Indeed block bots.

**Bulk operations**
- Bulk import from pasted CSV with optional rate-limited bulk-scoring after import.
- Audit log. Every Claude call writes a row with token counts and dollar cost. Cache hit rate and monthly spend are visible in the sidebar footer.

**Mobile**
- PWA web app for one-handed use from the phone. Quick log a new application from a recruiter call. Swipe right to mark applied, left to mark rejected. Tap to score. Daily briefing on the home tab.

## Architecture

```
Google Sheet (3 tabs)              Apps Script project
+----------------+                 +-------------------------+
| Applications   |  <-- writes --  | Code.gs (orchestration) |
| Analytics      |                 | Claude.gs (API client)  |
| Activity (log) |                 | Resume.gs (Drive cache) |
+----------------+                 | ResumeGen.gs            |
        ^                          | CoverLetter.gs          |
        |                          | InterviewPrep.gs        |
        | reads/writes             | Scraper.gs              |
        v                          | Gmail.gs                |
+----------------+                 | Drive.gs / Calendar.gs  |
| Gmail label    | --> trigger --> | Activity.gs             |
| Drive folders  |                 | BulkImport.gs           |
| Calendar       |                 | Onboarding.gs           |
| Drive Docs     | <-- reads ------+ Setup.gs                |
| (resume, gen)  |                 +-------------------------+
+----------------+                            |
                                              v
+--------------------------------+   UrlFetchApp + Web App
| Sidebar.html (in-Sheet UX)     |   |
| MobileApp.html (PWA)           |   v
| OnboardingDialog.html (modal)  | +-------------------------+
| BulkImportDialog.html (modal)  | | api.anthropic.com       |
+--------------------------------+ | claude-opus-4-7         |
                                   | adaptive thinking       |
                                   | prompt caching enabled  |
                                   +-------------------------+
```

## Prompt caching

The resume body is the cache anchor. It lives in a Drive Doc and is pulled via DocumentApp on every call, then snapshotted in CacheService for 6 hours. As long as the doc isn't edited, the bytes are identical across calls and the prefix cache hits at ~10% of input cost.

When the resume doc is edited, run "Refresh resume cache" from the menu. The next call pays full input cost; every call after hits cache normally.

Verify cache behavior in the hidden Activity tab: the `Cache Read Tokens` column should equal the resume + rubric token count on every call after the first.

## Install (manual, from this repo)

If you don't want to use the make-a-copy link above and would rather paste the code yourself, or you want to fork and modify:

1. Create a fresh Google Sheet
2. Extensions → Apps Script
3. Paste each `.gs` file from this repo into the editor (one Script file per `.gs`). Paste the HTML files (`Sidebar.html`, `MobileApp.html`, `OnboardingDialog.html`, `BulkImportDialog.html`, `LinkedInDialog.html`) as new HTML files (File → New → HTML).
4. Open `appsscript.json` in the editor (View → Show manifest file) and replace it with the version from this repo.
5. Save the project.
6. Run the `setup` function from the editor (or use the `Onboarding wizard` menu item from the Sheet for a richer flow). Authorize the OAuth scopes when prompted.
7. The wizard will ask for:
   - Anthropic API key (get one at console.anthropic.com)
   - Master resume Google Doc URL
   - Drive parent folder URL (where per-company sub-folders go)
   - Gmail label name (default `job-apply`)
   - Monthly spend threshold (default `20` USD)
   - Slack webhook URL, nudge target (`email`/`slack`/`both`), verification token, team_id allowlist (all optional, leave blank to skip)
   - OpenAI API key (optional, only needed for similar-applications search)
8. Reload the Sheet. The "Job CRM" menu appears.
9. Click "Open sidebar" to start.

### Slack deployment (optional)

If you set `SLACK_NUDGE_TARGET` to `slack` or `both`, the daily digest will POST to your incoming webhook URL.

If you also want the `/jobcrm` slash command:

1. Create a Slack app at api.slack.com → Slash Commands → New Command. URL is the Web App URL from step 12. Method: POST.
2. Web App access must be `Anyone` (Slack POSTs as an unauthenticated request). The script enforces a verification token + `team_id` allowlist instead of Google auth.
3. Copy the Verification Token from the Slack app's Basic Information page into `SLACK_VERIFICATION_TOKEN` (Apps Script → Project Settings → Script Properties).
4. Find your workspace `team_id` (starts with `T...`) and put it in `SLACK_TEAM_ID`.
5. Test from any Slack channel: `/jobcrm log Stripe | Senior Frontend Engineer | https://...`

### Mobile deployment

10. In the Apps Script editor: Deploy → New deployment
11. Type: Web app. Description: "Job CRM mobile". Execute as: User accessing the web app. Who has access: Only myself.
12. Deploy. Copy the resulting URL.
13. Open the URL on your phone in Safari (iOS) or Chrome (Android).
14. iOS: Share → Add to Home Screen. Chrome: menu → Install app.
15. The PWA opens full-screen and behaves like a native app.

The Web App URL is single-user. Only the deploying Google account can use it. Don't share the URL.

### Slack slash command (separate deployment)

The Slack slash command needs an `Anyone` access setting because Slack POSTs without Google auth. Create a second deployment from the same project:

1. Deploy → New deployment → Web app
2. Description: "Job CRM Slack endpoint". Execute as: User deploying the web app. Who has access: **Anyone**.
3. Deploy. Use this URL as the Slack slash command Request URL.

`doPost` distinguishes Slack requests from mobile-app requests by checking for `e.parameter.command` and `e.parameter.token`, then enforces a verification token + workspace `team_id` allowlist before doing anything. The mobile PWA URL stays single-user.

### Local development with clasp (optional)

If you want to edit code in your editor and push to Apps Script instead of pasting through the web IDE:

1. `npm install -g @google/clasp`
2. `clasp login`
3. Copy `.clasp.json.example` to `.clasp.json` and paste your Apps Script project ID (find it in Project Settings → IDs).
4. `clasp push` to upload local changes. `clasp pull` to sync changes made in the web IDE.
5. `.clasp.json` is gitignored so your project ID never lands in the repo.

## Publish as a template (for the project owner)

If you want to share your Sheet as a one-click template for others:

1. Open your bound Sheet.
2. File → Share → Share with others → "Anyone with the link, Viewer". (Edit access is not needed; copying creates an independent file.)
3. Copy the share URL. It looks like `https://docs.google.com/spreadsheets/d/<SHEET_ID>/edit`.
4. Replace the trailing `/edit` with `/copy` to create the make-a-copy URL.
5. Replace `REPLACE_WITH_TEMPLATE_SHEET_ID` near the top of this README with your actual `<SHEET_ID>`.

The copy includes the bound script. Each copier authorizes the OAuth scopes themselves and runs `setup` against their own Sheet, so there is no shared state and your API key stays in your project.

## Day 1 vs Day 2 split

**Day 1 (MVP, must-haves):**
- Setup, Resume cache, Activity log, Claude client, Code orchestration, Sidebar, Drive folders, Calendar events, JD scraper, Gmail import + daily nudge

**Day 2 (polish):**
- Tailored resume + cover letter generation with anti-fabrication rules
- Interview prep mode in the sidebar
- Mobile Web App / PWA
- Onboarding wizard
- Bulk CSV import

**Phase 3 (shipped):**
- Slack daily nudge (incoming webhook) + `/jobcrm log` slash command
- Vector embeddings + similar-applications search (OpenAI `text-embedding-3-small`)
- LinkedIn profile import for first-time resume generation

## Known limitations

- LinkedIn and Indeed scraping fails. Both serve Cloudflare-gated pages to non-browser fetches. Paste the JD body directly into the JD Text column for those.
- Apps Script execution limit is 6 minutes. A single score takes about 5-10 seconds. Bulk-score with rate limiting (3 sec per call) handles up to ~100 rows in a single execution.
- API key stored in Script Properties. Solo-use only. Production deployment for multiple users would require an OAuth proxy.
- No streaming. UrlFetchApp is synchronous; the JSON schema constraint keeps responses bounded.
- First call after a resume edit is a cache miss. Documented behavior; the menu has a refresh button to make this explicit.
- Mobile Web App can't do native push notifications or full offline mode (planned for Phase 3).
- Slack signature verification is not available. Apps Script `doPost(e)` exposes only `e.postData.contents` and `e.parameter`; HTTP headers (including `X-Slack-Signature`) are not surfaced. The script falls back to body-embedded verification token + `team_id` allowlist. Acceptable for personal single-workspace use; not suitable for a shared production deployment.
- Embedding generation calls OpenAI directly via UrlFetchApp. If you don't set `OPENAI_API_KEY`, scoring still works; the similar-applications feature is the only thing skipped.

## OAuth scopes

The first menu click triggers a Google consent screen for these scopes. Each is required by a specific feature. If you don't need a feature, you could remove the scope from `appsscript.json` before deploying, but the menu items that depend on it will break.

| Scope | Required by |
|---|---|
| `spreadsheets` | All sheet reads/writes (every menu item) |
| `gmail.modify` | Auto-import threads labeled `job-apply` and mark them processed |
| `gmail.send` | Daily nudge digest emails, cover-letter draft creation |
| `drive` | Per-company subfolder creation, tailored resume + cover letter Doc creation, master resume Doc reading |
| `documents` | Reading the master resume Doc body and writing tailored output Docs |
| `calendar` | Creating interview Calendar events on Status flip |
| `script.external_request` | Anthropic API + OpenAI embeddings + Slack webhook |
| `script.scriptapp` | Installing time-driven and onEdit triggers from `setup` |
| `userinfo.email` | First-run auth flow only (no persistent use) |
| `script.container.ui` | Sidebar, modals, and prompts inside the Sheet |

The scope list looks broad because the project actually does a lot. A reviewer reading `appsscript.json` will see exactly the scopes the code uses, no more.

## Privacy

- Sheet data, Drive folders, Activity log all live in my Google account. Nobody else can read them.
- Each scoring call sends JD text plus the resume body to the Anthropic API. Per the [Anthropic data use policy](https://www.anthropic.com/legal/commercial-terms), API data is not used for training.
- API key is in Script Properties: encrypted at rest, only readable by code in this script project.
- The `DRIVE_PARENT_FOLDER_ID` should point at a private folder you own. New per-company subfolders inherit the parent's sharing settings, so anything stored there (cover letters, tailored resumes) inherits whatever access the parent has. Use a folder that's NOT shared with others.
- The `Activity` tab is hidden by default. Even so, the `notes` column is sanitized before write: emails and phone numbers are redacted, and notes are capped at 500 characters.
- No third-party logging, no analytics SDK.
- This public repo excludes my actual API key, script project ID, and resume doc ID.

## Anti-aesthetic decisions

The sidebar and mobile app use the system font stack (no Inter, no Geist), a single deep-teal accent (`#2D5266`) against an off-white background (`#F8F8F6`), and a single 4px radius throughout. There are no purple-blue gradients, no `rounded-2xl` cards, no emoji headers. Skeleton loading instead of spinners. Empty states tell you what to do next instead of saying "No items yet". Solid bottom tab bar on mobile, not a hamburger drawer.

Cover angles, tailored resume bullets, and cover letter copy are constrained at the prompt level: no "Spearheaded", "Synergized", "Results-driven", "Excited to apply", or other AI-default vocabulary. Every angle has to reference something concrete from the candidate's experience that maps to a specific JD requirement.

The cover letter prompt requires referencing one verifiable specific about the company. If Claude can't ground in a real specific, it has to insert a `[VERIFY: ...]` placeholder rather than invent.

The tailored resume prompt forbids fabrication outright. If the JD asks for something not in the master resume, the output adds a `GAP: [requirement]` comment so the candidate sees the gap and decides, rather than seeing a tailored resume that quietly hallucinates skills.

## What I learned (case study notes)

To be filled in after a few weeks of real use:
- Applications tracked
- Cache hit rate
- Monthly Anthropic spend
- Average scoring latency
- Follow-ups automated
- Cover-letter drafts generated
- Tailored resumes generated
- What got cut, what surprised me, what I'd build next

## File map

| File | Purpose |
|---|---|
| `appsscript.json` | Manifest: V8 runtime, OAuth scopes, Web App config |
| `Setup.gs` | Idempotent installer, sheet creation, trigger setup |
| `Onboarding.gs` + `OnboardingDialog.html` | Richer first-run wizard with API key test |
| `Resume.gs` | Drive doc + CacheService snapshot for prompt-cache hits |
| `Activity.gs` | Audit log writer + cost calculator per model |
| `Claude.gs` | Raw HTTP client, scoring + extraction prompts, JSON schemas |
| `Code.gs` | Menu, scoring orchestration, sidebar handlers |
| `SidebarController.gs` | Sidebar quick action handlers, lifecycle |
| `Sidebar.html` | Primary in-Sheet UX, default + interview prep modes |
| `ResumeGen.gs` | Tailored resume generation with anti-fabrication rules |
| `CoverLetter.gs` | Cover letter drafting with anti-AI vocabulary bans |
| `InterviewPrep.gs` | Company research, likely questions, prep checklist |
| `Drive.gs` | Per-company folder management |
| `Calendar.gs` | onStatusEdit handler for interview events |
| `Scraper.gs` | JD URL fetch + HTML pre-clean + Claude extraction |
| `Gmail.gs` | Label parser (every 5min), daily follow-up nudge |
| `WebApp.gs` + `MobileApp.html` | Mobile PWA + Slack slash command router |
| `BulkImport.gs` + `BulkImportDialog.html` | CSV bulk import with optional rate-limited scoring |
| `Slack.gs` | Outbound nudge webhook + inbound slash command handler |
| `Embeddings.gs` | OpenAI embedding client, base64 Float32 storage, cosine similarity |
| `LinkedInImport.gs` + `LinkedInDialog.html` | One-shot resume generation from a LinkedIn profile export |
