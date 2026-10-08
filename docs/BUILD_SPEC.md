# BUILD SPEC — MHS Billing Tickets on Power Pages

> **For Claude:** this file is the complete brief. Build the app it describes as a **Power Pages code site** (a React + Vite + TypeScript single-page app) in VS Code, backed by Dataverse. Work through §10 *Build order* and check off §11 *Acceptance tests*. **§12 is a tested reference implementation** of every business rule. Port it as a module and don't re-derive the rules. Ask the owner before changing anything marked **LOCKED**.
>
> **For the owner:** in Claude Enterprise, start a Claude Code session on the repo **brendannyates/mhs-billing-excel-addin** (it contains a working vanilla-JS build of both editions, `pages/` + `rules.js`), attach this file, and say *"Build this."* **Never attach `migration-data.json` or any file with patient data.**

---

## 1. What we're building

An internal **patient billing escalation ticket desk** for Mindful Health Solutions (MHS). Clinic and operations staff log billing escalations through an existing **Microsoft Form**. Each escalation is emailed to the external billing vendor **Arietis** (`patientbilling@arietishealth.com`). Staff then track each ticket until Arietis confirms receipt and resolves it.

The site replaces an Excel-based tracker. It has to feel like a lightweight ticketing system:
- one-line ticket rows
- details only when a ticket is opened
- fast filtering by MRN
- SLA clocks
- a needs-action queue
- an ops-escalation review mode for meetings with Arietis

**Site:** `https://testbilling.powerappsportals.com` (Power Pages, Entra ID sign-in). **Audience:** MHS staff only (**LOCKED**). Arietis never signs in.

## 2. Architecture (LOCKED unless the owner agrees otherwise)

```
Microsoft Form (intake, unchanged) ──Power Automate──► Dataverse: mhs_ticket
billing@mindfulhealthsolutions.com inbox ──flow──► mhs_ticket (Arietis reply times)
send-to-Arietis flow ──► mhs_ticket (Outlook Conversation ID)
Daily/weekly digest flow ──► Office Script (shared rules) ──► personalized emails
Power Pages code site (this app) ◄──► Dataverse via Power Pages Web API (/_api/...)
```

- **Frontend:** a Power Pages **code site** (single-page app), uploaded with `pac pages upload-code-site`. Use React 18, Vite and TypeScript. Plain CSS with variables, no UI framework, keeps the bundle small; Fluent UI React v9 is acceptable if preferred.
- **Data:** Dataverse tables with the publisher prefix **`mhs_`**, in the solution **BillingEscalations**.
- **Auth:** Power Pages Entra ID sign-in. The site stays **Private**, so only signed-in staff can reach it. The current user comes from `window.Microsoft?.Dynamic365?.Portal?.User`.
- **Licensing guardrails (LOCKED):** standard connectors only, so no premium connectors and no Azure resources. No custom APIs. Power Automate uses the Microsoft Forms, Dataverse, Office 365 Outlook and Excel Online (Business) "Run script" connectors.
- **PHI:** patient code and MRN are PHI-adjacent. Never log them to the console, never put them in URLs, never send them to third parties. Vendor exports contain only the vendor-safe columns listed in §6.7.

## 3. Data model (Dataverse)

All tables are organization-owned, with the primary column `mhs_name`. Create them with `scripts/dataverse-setup.mjs` in the repo, or by hand with exactly these names. Date/time columns use **User local** behavior; date columns use **Date only**.

### 3.1 `mhs_ticket` (set `mhs_tickets`). Alternate key: `mhs_formid`. `mhs_name` = `"#<formid> <patient>"`
| Logical name | Type | Display name | Source (Form question / origin) |
|---|---|---|---|
| `mhs_formid` | int | Ticket # | Id |
| `mhs_starttime` | dt | Form start time | Start time |
| `mhs_submittedon` | dt | Submitted on | Completion time |
| `mhs_owneremail` | text | Owner email | Email |
| `mhs_ownername` | text | Owner name | Name |
| `mhs_requesttype` | text | Request type | Choose an option: |
| `mhs_department` | text | Department | Department Submitting Ticket: |
| `mhs_clinic` | text | Clinic | Patient Clinic: |
| `mhs_patient` | text | Patient | Patient: |
| `mhs_mrn` | text | MRN | MRN: |
| `mhs_urgency` | text | Urgency | Urgency Level: |
| `mhs_source` | text | Source of inquiry | Source of Inquiry: |
| `mhs_tasktype` | text | Task type | Task Type: |
| `mhs_amount` | num | Amount | Amount: |
| `mhs_notes` | memo | Notes | Notes / Encounter# / Other relevant addition information: |
| `mhs_attachments` | memo | Attachments | Please add any attachments here: |
| `mhs_status` | text | Status | Status: |
| `mhs_firstreplydate` | date | Date of first reply from Arietis | Date of first reply from Arietis: |
| `mhs_outreachdate` | date | Date of patient outreach | Date of Patient Outreach (if applicable): |
| `mhs_resolutiondate` | date | Date of resolution | Date of Resolution: |
| `mhs_falseverification` | text | False verification by Arietis | Was false verification of balance/charge given by Arietis? |
| `mhs_servicerecovery` | text | Service recovery flag | Patient Feedback/Service Recovery Flag: |
| `mhs_errorsource` | text | Source of error | Source of Error: |
| `mhs_outcome` | text | Outcome | Outcome: |
| `mhs_ehrtask` | bool | EHR task | EHR Task |
| `mhs_formlink` | memo | Form response link | Form Response Link |
| `mhs_conversationid` | text | Arietis conversation ID | ArietisConversationId |
| `mhs_senttoarietisat` | dt | Sent to Arietis at | SubmittedToArietisAt |
| `mhs_firstreplyat` | dt | Arietis first reply at | ArietisFirstReplyAt |
| `mhs_lastreplyat` | dt | Arietis last reply at | ArietisLastReplyAt |
| `mhs_replycount` | int | Arietis reply count | ArietisReplyCount |
| `mhs_cc` | text | CC | CC |
| `mhs_lastupdatedby` | text | Last updated by | LastUpdatedBy |
| `mhs_lastupdatedat` | dt | Last updated at | LastUpdatedAt |
| `mhs_formsyncedat` | dt | Form synced at | RawSyncedAt |
| `mhs_worknotes` | memo | Work notes | WorkNotes |
| `mhs_opsescalation` | text | Ops escalation | OpsEscalation |
| `mhs_opsreason` | text | Ops escalation reason | OpsReason |
| `mhs_opsescalatedby` | text | Ops escalated by | OpsEscalatedBy |
| `mhs_opsescalatedat` | dt | Ops escalated at | OpsEscalatedAt |
| `mhs_arietiscommitment` | memo | Arietis commitment | ArietisCommitment |
| `mhs_followupdue` | date | Follow-up due | FollowUpDue |
| `mhs_reviewnotes` | memo | Review notes | ReviewNotes |
| `mhs_lastreviewedat` | dt | Last reviewed at | LastReviewedAt |

### 3.2 `mhs_ticketactivity` (set `mhs_ticketactivities`): the audit log. One row per changed field.
| Logical name | Type | Display name |
|---|---|---|
| `mhs_at` | dt | When |
| `mhs_ticketnumber` | int | Ticket # |
| `mhs_name` | text | Field |
| `mhs_oldvalue` | memo | Old value |
| `mhs_newvalue` | memo | New value |
| `mhs_by` | text | By |
| `mhs_source` | text | Source |

### 3.3 Lookup tables
- **`mhs_setting`** (`mhs_settings`): `mhs_name`, `mhs_value`, `mhs_notes`. Keys:
  - `ReceiptSLABusinessDays`=2
  - `ResolutionSLABusinessDays`=5
  - `StaleAfterBusinessDays`=3
  - `ArietisEmail`=patientbilling@arietishealth.com
  - `BillingInbox`=billing@mindfulhealthsolutions.com
  - `LeadershipRecipients` (semicolon-separated)
  - `ClinicDailyDigest`=No
  - `ArietisDailyFollowup`=No
  - `WorkbookUrl`
  - `FormUrl`=https://forms.cloud.microsoft/Pages/ResponsePage.aspx?id=8Sq2y8CkOEWdeGKInLNdYw1_rrxVxaFImLkzEWnt0GFUQUtJVFg2VzU5UEowM01JMUgzUUpKMTZPVi4u
- **`mhs_holiday`** (`mhs_holidays`): `mhs_name`, `mhs_date` (date only). Skipped by every business-day calculation.
- **`mhs_refoption`** (`mhs_refoptions`): `mhs_name` (the option text), `mhs_category`, `mhs_email`, `mhs_sort`. Categories, exactly as they come from the old REF sheet:
  - `ClinicName` (21 clinics, with each clinic's help inbox in `mhs_email`)
  - `Department Submitting Ticket:`
  - `Urgency Level:`
  - `Source of Inquiry`
  - `Task Type`
  - `Status`
  - `Was false verification of balance/charge given by Arietis?`
  - `Patient Feedback/Service Recovery Flag:`
  - `Source of Error:`
  - `Outcome:`

**Every dropdown in the app reads its options from `mhs_refoption`.** Never hard-code option lists, except as fallbacks.

### 3.4 Status values (from REF)
- Open:
  - `MHS - Submitted to Arietis` → stage **Awaiting receipt**
  - `Arietis - Confirmed Receipt` → **With Arietis**
  - `Arietis - Confirmed Receipt - Pending Patient Call` → **Pending call**
- Closed:
  - `Closed (No Corrective Action Needed)` → **Closed**
  - `Resolved (Corrective Action Complete)` → **Closed**

## 4. Business rules (implemented in §12, which you port as-is)

1. **Business days** skip weekends and the holidays in `mhs_holiday`. Due = the submission *day* + N business days, through end of day. A ticket submitted on a weekend counts from the next business day. Example: submitted Thursday → receipt due Monday, resolution due the following Thursday.
2. **Receipt SLA (2 BD)** stops at the first Arietis reply (`mhs_firstreplyat`), else at `mhs_firstreplydate`. If neither exists but the status has moved past *Awaiting receipt*, it is "met (no date)".
   - States: `met`, `late`, `met-nodate`, `pending`, `due-today`, `breached` (with business days over).
3. **Resolution SLA (5 BD)** stops when the status is Closed or Resolved, at `mhs_resolutiondate`, else at `mhs_lastupdatedat`. Same states as the receipt SLA.
4. **Needs-action rules.** Each produces an action with a severity (1 critical, 2 warning, 3 info) and who acts (owner, Arietis or ops):
   - Receipt overdue (1, Arietis)
   - Resolution overdue (1, Arietis)
   - Receipt due today (2)
   - Resolution due today (2)
   - Arietis replied but status still *Submitted* (2, owner)
   - New Arietis reply since the last update (2, owner)
   - Clinic or patient blank (2, owner)
   - *Pending call* with no outreach date (2, owner)
   - Receipt met but no date recorded (3, owner)
   - Stale: no activity for 3 or more BD (3, owner; suppressed when something is already overdue)
   - Closed but missing resolution date, source of error, outcome or false-verification (3, owner)
   - Ops follow-up overdue (1, Arietis)
   - Ops follow-up due today (2, Arietis)
   - Ops escalation not yet reviewed (3, ops)
5. **Ops escalation**
   - A ticket is escalated when any of these is true, unless `mhs_opsescalation = "Cleared"`:
     - `mhs_opsescalation = "Yes"` (a leader flagged it, with a reason)
     - urgency is Critical
     - service recovery = Yes
   - "Suggested" = open tickets past SLA that aren't escalated.
6. **Ownership:** **Mine** = tickets where `mhs_owneremail` equals the signed-in user's email (case-insensitive), plus tickets that list the user in `mhs_cc`.
7. **Edits** PATCH only the changed columns, and always set `mhs_lastupdatedby` and `mhs_lastupdatedat`. They also add one `mhs_ticketactivity` row per changed field, with `mhs_source = "Web"`. Multi-select fields (task type, outcome) are stored `;`-joined and compared as sets.
8. **Closing assist:** when the status changes to Closed or Resolved and the resolution date is blank, prefill it with today. When the status moves past *Submitted* and the first-reply date is blank, prefill it with the first-reply timestamp, or today. If resolution date, source of error, outcome or false-verification is missing, show a non-blocking hint: "To close cleanly, also fill: …".
9. **Times:** Dataverse returns UTC. Display in the browser's local time; convert with `utcIsoToSerial` / `serialToUtcIso` in §12.

## 5. Look and feel

- **Functional ticket-queue aesthetic.** Neutral surfaces (`#f4f5f7` background, white panels, `#e2e5e9` lines) and one accent (`#1f5f8b`).
- Color is used only for state:
  - **Status chips:** Awaiting = amber, With Arietis = blue, Pt call = violet, Closed = green.
  - **Severity dots:** 1 = red `#b42318`, 2 = amber `#9a5800`, 3 = gray.
- Segoe UI. Cascadia Mono or Consolas for ticket #, patient code and MRN. Tabular numerals throughout.
- Dense layout with 32–44 px rows, centered at max 1040 px. Dark mode via `prefers-color-scheme`. Visible keyboard focus, and WCAG AA contrast.
- **LOCKED: the collapsed ticket row is one line:** `● #71  PATIENT  MRN 100200  Issue (short)  [Status chip]`
  - The dot is the highest action severity, or nothing if no action is needed.
  - Short issue names:
    - Balance dispute
    - Duplicate balance
    - Refund
    - Callback request
    - Superbill / EOB
    - Can't reach Arietis
    - Not billed to insurance
  - "+N" when there are multiple task types.
  - Nothing else goes on the row: no amounts, owners, flags or SLA text.

## 6. Screens

Header: app name, last-refreshed time, a refresh button, the signed-in user's name, and a settings gear. Tabs:

### 6.1 Mine (default)
- Filter box: **"Filter by MRN, patient code or #"**. The same box appears on Mine, Action, Ops and Clinic. It matches MRN substrings, patient-code substrings and exact ticket #. Counts and sections recompute as you type.
- Segments: **Open · Needs action · Closed**, each with a count. Open is sorted by attention: severity, then urgency, then oldest first.

### 6.2 Action
- Everyone/Mine toggle and a clinic dropdown.
- Buttons:
  - **Email Arietis · N at/past SLA**: a bulk follow-up draft.
  - **Remind owners · N**: a sheet with one Compose button per owner.
- Sections: **Past SLA**, **Due today / update needed**, **Housekeeping**. While filtering, hide sections with no results.

### 6.3 Ops (escalation review)
- Segments: **Escalated · F/U due · Unreviewed · Suggest · Closed**.
- **Review with Arietis:** a full-screen, one-ticket-at-a-time, *share-safe* mode for screen-sharing on calls. It hides work notes, staff emails and escalation reasons, and shows a green "Share-safe view" banner.
  - Shows: #, patient, MRN, clinic, issue, amount, submitted date and business days open, status, both SLA clocks, last Arietis reply, and the form details.
  - Captures: Status, **Arietis commitment**, **Follow-up due**, **Review notes**. It stamps `mhs_lastreviewedat`.
  - Buttons: **Save & next**, **De-escalate** (sets `Cleared`), ←/→ and Exit.
- **Export for Arietis:** a table of vendor-safe columns only — Ticket #, Patient, MRN, Clinic, Issue, Amount, Submitted, Business days open, Status, Receipt SLA, Resolution SLA, Last Arietis reply, Issue details, Arietis commitment, Follow-up due.
  - Actions: **Email to Arietis**, which opens an Outlook draft to patientbilling@ with CC billing@, and **Copy table**, which copies HTML plus TSV.
  - A unit test must prove that internal notes and reasons are never exported.

### 6.4 Clinic
- Clinic dropdown, remembered per browser, and **Email clinic**, which drafts to that clinic's help inbox from `mhs_refoption.mhs_email`.
- Segments: **Open · Awaiting · Arietis · Pt call · Closed**. No KPI tiles here.

### 6.5 Stats
- Range filter: All / 90 days / 30 days / Month to date, by submission date.
- KPI tiles:
  - Open
  - Past SLA
  - Receipt ≤2 BD %
  - Resolved ≤5 BD %
  - Avg BD to receipt
  - Avg BD to resolve
  - $ open
  - Service recovery
- **Clinic table:** Open, Awaiting, With Arietis, Past SLA, Receipt %, Resolution %, Avg BD, Total, and a stage-mix bar. Clicking a row opens that clinic.
- Bar lists: task type, open by owner, source of error (closed), outcome (closed), source of inquiry, flags.

### 6.6 MRN lookup
Search by MRN, patient code or #. Results are grouped by patient: "MRN 100200 · ABCDEF — 3 tickets · 1 open · $1,234".

### 6.7 + New
- Embeds the Form in an iframe (`FormUrl` + `&embed=true`), with **Open in new tab**.
- Note: "The new ticket appears here within about a minute."

### 6.8 Ticket detail (opens from any row)
- Header: #, stage chip, flags, age in BD.
- SLA timeline: submitted; receipt due → state; resolution due → state.
- Needs-action list, and buttons **Email Arietis** (subject `[BE-<id>] Status request — patient <code> (MRN …)`) and **Email owner**.
- **Ops escalation card:** shows the reasons, commitment, follow-up and review notes, plus De-escalate. If the ticket isn't escalated, it shows **Escalate to ops** with a reason field.
- **Update form:**
  - Status
  - First reply date, outreach date, resolution date
  - Urgency, source of error, false verification, service recovery
  - EHR task
  - Outcome (multi-select)
  - CC
  - Work notes
- **Correct escalation details** (collapsed): clinic, department, patient, MRN, source, amount, task types, notes.
- Arietis thread: sent, first and last reply, count, conversation ID.
- History from `mhs_ticketactivity`.
- Sticky save bar: "N unsaved changes", Revert, **Save** (Ctrl+Enter). Leaving with unsaved changes needs a second press. Don't use `window.confirm`.

**Emails from the app** open an Outlook web compose deeplink (`https://outlook.office.com/mail/deeplink/compose?to=&cc=&subject=&body=`). The formatted HTML version is also copied to the clipboard; if the URL would exceed about 1,800 characters, the body says "press Ctrl+V".

**Auto-refresh** every 60 s while the tab is visible. Never overwrite a form that has unsaved changes.

## 7. Power Pages configuration
1. **Table permissions**, Global access, for the *Authenticated Users* role (or a "Billing Staff" role):
   - `mhs_ticket`: Read and Write
   - `mhs_ticketactivity`: Read and Create
   - `mhs_setting`, `mhs_holiday`, `mhs_refoption`: Read
2. **Site settings:**
   - `Webapi/<table>/enabled = true` for each of the 5 tables.
   - `Webapi/<table>/fields` = an explicit column list, never `*`. For `mhs_ticket` it is: `mhs_ticketid,mhs_formid,mhs_starttime,mhs_submittedon,mhs_owneremail,mhs_ownername,mhs_requesttype,mhs_department,mhs_clinic,mhs_patient,mhs_mrn,mhs_urgency,mhs_source,mhs_tasktype,mhs_amount,mhs_notes,mhs_attachments,mhs_status,mhs_firstreplydate,mhs_outreachdate,mhs_resolutiondate,mhs_falseverification,mhs_servicerecovery,mhs_errorsource,mhs_outcome,mhs_ehrtask,mhs_formlink,mhs_conversationid,mhs_senttoarietisat,mhs_firstreplyat,mhs_lastreplyat,mhs_replycount,mhs_cc,mhs_lastupdatedby,mhs_lastupdatedat,mhs_formsyncedat,mhs_worknotes,mhs_opsescalation,mhs_opsreason,mhs_opsescalatedby,mhs_opsescalatedat,mhs_arietiscommitment,mhs_followupdue,mhs_reviewnotes,mhs_lastreviewedat`.
3. **Requests:** use the Power Pages Web API at `/_api/<entityset>`, with `credentials: "same-origin"`. Every write sends the header `__RequestVerificationToken`, from `window.shell.getTokenDeferred()` or, as a fallback, `GET /_layout/tokenhtml`. Follow `@odata.nextLink`, and send the header `Prefer: odata.maxpagesize=5000`.
4. **Blocked attachments:** if the upload rejects `.js` files, remove `js` from the environment's blocked attachments.
5. **Deploy:** `npm run build`, then `pac pages upload-code-site --rootPath . --compiledPath ./dist --siteName "Billing Escalations"`. Then reactivate the site in Power Pages → *Inactive sites*.

## 8. Power Automate flows

Every flow uses standard connectors, with trigger concurrency set to 1.

1. **Intake.** Forms *When a new response is submitted* → *Get response details* → Dataverse *List rows* on `mhs_formid eq <responseId>`.
   - If no row exists: *Add a new row*, mapping every Form answer to its column (§3.1).
   - If a row exists (an edit made through the Form's edit link): *Update a row*, using `if(empty(x), null, x)` for each column so that **non-blank answers merge and nothing is ever blanked**, plus an activity row.
2. **Send-to-Arietis.** In the existing flow, after the email is sent, set `mhs_conversationid` and `mhs_senttoarietisat`. Prefix the subject with `[BE-<id>]`.
3. **Reply tracker.** Trigger: *When a new email arrives in a shared mailbox (V2)* on `billing@mindfulhealthsolutions.com`, From `patientbilling@arietishealth.com`, with a matching trigger condition.
   - Match the ticket by `mhs_conversationid`, falling back to `[BE-n]` in the subject.
   - Skip the ticket if its status starts with Closed or Resolved.
   - Update the ticket:
     - set `mhs_firstreplyat` if blank
     - set `mhs_firstreplydate` if blank
     - set `mhs_lastreplyat`
     - `mhs_replycount` + 1
   - Add an activity row.
4. **Daily digests** (weekdays 08:00 and 17:00 Pacific) and **weekly** (Friday 16:00).
   - List the 5 tables, then *Excel Online (Business) › Run script* `DigestsDv` (office-scripts/DigestsDv.ts in the repo) in a blank helper workbook, passing each table as JSON text.
   - Send each returned `{to, cc, subject, html}` with *Send an email from a shared mailbox (V2)* from billing@.
   - Content:
     - **Start of day:** needs action, past SLA, due today, open tickets.
     - **End of day:** today's Arietis replies, today's changes, due the next business day.
     - **Weekly:** opened, closed, on-time %, aging, escalation flags, closed this week.
     - **Leadership:** a weekly clinic rollup, plus daily past-SLA flags with an Ops escalations section.

## 9. Data migration
Run `node scripts/dataverse-setup.mjs --env <org url> --data migration-data.json`, from the repo. It creates the tables, seeds settings, holidays and REF options, and imports the existing tickets idempotently, skipping existing `mhs_formid` values. The owner runs it locally because the data file contains PHI. Never commit it or upload it to any AI tool.

## 10. Build order
1. Scaffold the Vite + React + TS app. Add `src/lib/rules.ts` = §12 (`src/rules.ts` in the repo), **verbatim**, with `export` added to each top-level declaration. Port `tests/rules.test.mjs` first.
2. Write `src/lib/dataverse.ts` with typed `getAll`, `patch` and `post`, the token helper and the user helper. Map records to the core's row shape with `dvTicketsToTable` and `dvActivityRows`.
3. Build the app shell: header, tabs, filter box, the one-line row component, the detail view and the save bar.
4. Build the screens in this order: Mine → Detail → Action → Clinic → MRN → Stats → Ops/Review/Export → New.
5. Add email composition (deeplink and clipboard).
6. Add a local dev mock of `/_api`, a tiny Vite middleware with synthetic data, so the app runs with `npm run dev` without Power Pages.
7. Do the Power Pages config (§7), the upload, and the flows (§8).

## 11. Acceptance tests (must pass)
- Business days: Mon→Wed (2 BD); Thu 4 pm→Mon (2 BD); Sat→Tue (2 BD); Fri 9/4/2026→Wed 9/9 (2 BD, Labor Day skipped); Wed 11/25/2026→Tue 12/1 (2 BD, Thanksgiving and the day after skipped).
- A resolved ticket never logs Arietis replies. The first reply sets first and last; later replies move only last; the count increments.
- Ops queue: Critical, service recovery and manual flags are included. A ticket with `Cleared` is excluded even when Critical. An overdue follow-up raises a severity-1 action.
- The vendor export never contains work notes, review notes, escalation reasons or staff emails.
- Saving changes only the edited columns, and writes one activity row per changed field. Reordering multi-select checkboxes is *not* a change.
- The MRN filter narrows every list and its counts. Clearing it restores them.
- The one-line row never wraps at 360 px width. Keyboard: `/` jumps to lookup, Esc goes back, Ctrl+Enter saves.
- No console errors. No PHI in URLs or console output.

## 12. Reference implementation — `src/rules.ts` (port verbatim)
This is the tested rules engine from the repo (`src/rules.ts`, 29 passing tests including `tests/rules.test.mjs`). It's TypeScript compatible with Office Scripts, so it has no imports, no `any` and no DOM. In the React app, export the functions from it. The `COL` names are the Excel Master headers; keep them, because the Dataverse mapping (`DV_TICKET_FIELDS`) is defined in terms of them.

```ts
// ============================================================================
// MHS Billing Tickets — rules engine (single source of truth)
// SLA clocks, needs-action rules, ops escalation, stats and email builders.
// Compiled to rules.js (browser + Power Pages) and inlined into Office Scripts by scripts/build.mjs.
// Must stay Office-Scripts compatible: no `any`, no imports, no DOM.
// All date/times are Excel serial numbers in LOCAL wall-clock time.
// ============================================================================

type Cell = string | number | boolean;

const COL = {
  id: "Id",
  start: "Start time",
  submitted: "Completion time",
  email: "Email",
  name: "Name",
  option: "Choose an option:",
  dept: "Department Submitting Ticket:",
  clinic: "Patient Clinic:",
  patient: "Patient:",
  mrn: "MRN:",
  urgency: "Urgency Level:",
  source: "Source of Inquiry:",
  taskType: "Task Type:",
  amount: "Amount:",
  notes: "Notes / Encounter# / Other relevant addition information:",
  attachments: "Please add any attachments here:",
  status: "Status:",
  firstReplyDate: "Date of first reply from Arietis:",
  outreachDate: "Date of Patient Outreach (if applicable):",
  resolutionDate: "Date of Resolution:",
  falseVerif: "Was false verification of balance/charge given by Arietis?",
  serviceRecovery: "Patient Feedback/Service Recovery Flag:",
  errorSource: "Source of Error:",
  outcome: "Outcome:",
  // ---- operational columns on Master (added by Initialize; never present in Raw Data) ----
  ownerEmail: "Owner Email",             // reassignable recipient; defaults to the Forms Email
  convId: "Vendor Conversation ID",
  sentAt: "Submitted to Vendor At",
  firstReplyAt: "First Reply At",        // flows may instead write the timestamp into "Date of first reply from Arietis:"
  lastReplyAt: "Last Reply At",
  closedAt: "Closed At",
  confirmDue: "Confirmation Due",
  resolveDue: "Resolution Due",
  escFlag: "Escalation Flag",
  conflicts: "Sync Conflicts",
  followUpAt: "Follow Up At",
  reminderFlags: "Reminder Flags",
  ehr: "EHR Task",
  link: "Form Response Link",
  replyCount: "Arietis Reply Count",
  cc: "CC",
  updBy: "Last Updated By",
  updAt: "Last Updated At",
  rawSynced: "Raw Synced At",
  workNotes: "Work Notes",
  opsFlag: "Ops Escalation",             // "Yes" = escalated by a leader, "Cleared" = de-escalated, blank = auto rules
  opsReason: "Ops Reason",
  opsBy: "Ops Escalated By",
  opsAt: "Ops Escalated At",
  commitment: "Arietis Commitment",
  followUpDue: "Follow-up Due",
  reviewNotes: "Review Notes",
  reviewedAt: "Last Reviewed At",
};

/** The 24 Microsoft Forms columns, in Raw Data / Master order. */
const FORM_COLS: string[] = [
  COL.id, COL.start, COL.submitted, COL.email, COL.name, COL.option, COL.dept, COL.clinic,
  COL.patient, COL.mrn, COL.urgency, COL.source, COL.taskType, COL.amount, COL.notes,
  COL.attachments, COL.status, COL.firstReplyDate, COL.outreachDate, COL.resolutionDate,
  COL.falseVerif, COL.serviceRecovery, COL.errorSource, COL.outcome,
];
/** Operational columns appended to Master by Initialize (order = order added). */
const EXTRA_COLS: string[] = [
  COL.ownerEmail, COL.convId, COL.sentAt, COL.lastReplyAt, COL.closedAt, COL.confirmDue, COL.resolveDue, COL.escFlag,
  COL.conflicts, COL.followUpAt, COL.reminderFlags, COL.ehr, COL.cc, COL.updBy, COL.updAt,
  COL.workNotes, COL.opsFlag, COL.opsReason, COL.opsBy, COL.opsAt, COL.commitment, COL.followUpDue, COL.reviewNotes, COL.reviewedAt,
];
const MASTER_COLS: string[] = FORM_COLS.concat(EXTRA_COLS);
const DATE_COLS: string[] = [
  COL.start, COL.submitted, COL.firstReplyDate, COL.outreachDate, COL.resolutionDate, COL.sentAt, COL.firstReplyAt,
  COL.lastReplyAt, COL.closedAt, COL.followUpAt, COL.updAt, COL.rawSynced, COL.opsAt, COL.followUpDue, COL.reviewedAt,
];
/** Date-only columns (written as dates, not timestamps). */
const DATE_ONLY_COLS: string[] = [COL.firstReplyDate, COL.outreachDate, COL.resolutionDate, COL.followUpDue, COL.followUpAt];
/** Fields staff own on Master (Portal rule): the form only fills them while blank and they never raise sync conflicts. */
const MASTER_OWNED_COLS: string[] = [
  "Status:", "Date of first reply from Arietis:", "Date of Patient Outreach (if applicable):", "Date of Resolution:",
  "Was false verification of balance/charge given by Arietis?", "Patient Feedback/Service Recovery Flag:", "Source of Error:", "Outcome:", "EHR Task",
];
/** Form fields that are never merged after import ("Choose an option:" flips to "Update…" when owners use the edit link). */
const NEVER_MERGE_COLS: string[] = ["Choose an option:", "Form Edit Submission Link", "Column1"];
/** Form fields a Forms "edit response" may legitimately change on an existing ticket. */
const MERGEABLE_COLS: string[] = [
  COL.dept, COL.clinic, COL.patient, COL.mrn, COL.urgency, COL.source, COL.taskType, COL.amount,
  COL.notes, COL.attachments, COL.status, COL.firstReplyDate, COL.outreachDate, COL.resolutionDate,
  COL.falseVerif, COL.serviceRecovery, COL.errorSource, COL.outcome,
];

const STATUS = {
  submitted: "MHS - Submitted to Arietis",
  confirmed: "Arietis - Confirmed Receipt",
  pendingCall: "Arietis - Confirmed Receipt - Pending Patient Call",
  closedNoAction: "Closed (No Corrective Action Needed)",
  resolved: "Resolved (Corrective Action Complete)",
};
const CLOSED_STATUSES: string[] = [STATUS.closedNoAction, STATUS.resolved];

const DEFAULTS = {
  receiptDays: 2,
  resolutionDays: 5,
  staleDays: 3,
  arietisEmail: "patientbilling@arietishealth.com",
  billingInbox: "billing@mindfulhealthsolutions.com",
};

// ---------------------------------------------------------------------------
// Value helpers
// ---------------------------------------------------------------------------
function norm(v: Cell | null | undefined): string {
  if (v === null || v === undefined) return "";
  return String(v).replace(/ /g, " ").replace(/\s+/g, " ").trim();
}
function normKey(v: Cell | null | undefined): string {
  return norm(v).toLowerCase();
}
function isBlank(v: Cell | null | undefined): boolean {
  return v === null || v === undefined || norm(v) === "";
}
function serialFromParts(y: number, mo: number, d: number, h: number, mi: number, s: number): number {
  return Date.UTC(y, mo - 1, d, h, mi, s) / 86400000 + 25569;
}
/** Wall-clock offset used for ISO timestamps with Z/offset. Browser sets local; Office Scripts use Pacific. */
let OFFSET_FN: (ms: number) => number = (ms: number): number => pacificOffsetMin(ms);
function setOffsetFn(f: (ms: number) => number): void { OFFSET_FN = f; }
/** Parses an Excel serial, ISO string, or US m/d/yy[ h:mm AM] string to a serial. */
function toSerial(v: Cell | null | undefined): number | null {
  if (v === null || v === undefined || v === "" || typeof v === "boolean") return null;
  if (typeof v === "number") return v > 0 ? v : null;
  const s = norm(v);
  if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}(:\d{2}(\.\d+)?)?(Z|[+-]\d{2}:?\d{2})$/i.test(s)) {
    const ms = Date.parse(s);
    if (!isNaN(ms)) return (ms + OFFSET_FN(ms) * 60000) / 86400000 + 25569;
  }
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T ](\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (m) {
    return serialFromParts(+m[1], +m[2], +m[3], m[4] ? +m[4] : 0, m[5] ? +m[5] : 0, m[6] ? +m[6] : 0);
  }
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AaPp][Mm])?)?$/);
  if (m) {
    let y = +m[3];
    if (y < 100) y += 2000;
    let h = m[4] ? +m[4] : 0;
    const ap = m[7] ? m[7].toUpperCase() : "";
    if (ap === "PM" && h < 12) h += 12;
    if (ap === "AM" && h === 12) h = 0;
    return serialFromParts(y, +m[1], +m[2], h, m[5] ? +m[5] : 0, m[6] ? +m[6] : 0);
  }
  return null;
}
function serialToDate(s: number): Date {
  return new Date(Math.round((s - 25569) * 86400000));
}
function pad2(n: number): string {
  return n < 10 ? "0" + n : String(n);
}
function fmtDate(s: number | null): string {
  if (s === null) return "";
  const d = serialToDate(s);
  return (d.getUTCMonth() + 1) + "/" + d.getUTCDate() + "/" + String(d.getUTCFullYear()).slice(2);
}
function fmtShort(s: number | null): string {
  if (s === null) return "";
  const d = serialToDate(s);
  return (d.getUTCMonth() + 1) + "/" + d.getUTCDate();
}
function fmtTime(s: number | null): string {
  if (s === null) return "";
  const d = serialToDate(s);
  let h = d.getUTCHours();
  const ap = h >= 12 ? "pm" : "am";
  h = h % 12;
  if (h === 0) h = 12;
  return h + ":" + pad2(d.getUTCMinutes()) + ap;
}
function fmtDateTime(s: number | null): string {
  if (s === null) return "";
  const hasTime = Math.abs(s - Math.floor(s)) > 1e-6;
  return hasTime ? fmtShort(s) + " " + fmtTime(s) : fmtShort(s);
}
function fmtMoney(n: number | null): string {
  if (n === null || isNaN(n)) return "";
  return "$" + Math.round(n).toLocaleString("en-US");
}
function weekdayName(s: number): string {
  return ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"][weekday(s)];
}
function escHtml(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

// ---------------------------------------------------------------------------
// Business-day math (weekends + holiday table)
// ---------------------------------------------------------------------------
function weekday(s: number): number {
  return (((Math.floor(s) - 25569 + 4) % 7) + 7) % 7; // 0 = Sunday
}
function isBizDay(day: number, hol: Set<number>): boolean {
  const w = weekday(day);
  return w !== 0 && w !== 6 && !hol.has(Math.floor(day));
}
/** Date (whole-day serial) that is n business days after the day of s. */
function addBizDays(s: number, n: number, hol: Set<number>): number {
  let d = Math.floor(s);
  let c = 0;
  while (c < n) {
    d += 1;
    if (isBizDay(d, hol)) c += 1;
  }
  return d;
}
/** Business days elapsed from the day of a to the day of b (a's day excluded). Negative if b < a. */
function bizDaysBetween(a: number, b: number, hol: Set<number>): number {
  const da = Math.floor(a);
  const db = Math.floor(b);
  if (db === da) return 0;
  const sign = db > da ? 1 : -1;
  const lo = Math.min(da, db);
  const hi = Math.max(da, db);
  let c = 0;
  for (let d = lo + 1; d <= hi; d++) if (isBizDay(d, hol)) c += 1;
  return sign * c;
}

// ---------------------------------------------------------------------------
// Column index
// ---------------------------------------------------------------------------
interface ColIndex {
  map: Map<string, number>;
}
function buildIndex(headers: Cell[]): ColIndex {
  const map = new Map<string, number>();
  for (let i = 0; i < headers.length; i++) {
    const k = normKey(headers[i]);
    if (k && !map.has(k)) map.set(k, i);
  }
  return { map: map };
}
function ci(idx: ColIndex, name: string): number {
  const v = idx.map.get(normKey(name));
  return v === undefined ? -1 : v;
}
function cell(row: Cell[], idx: ColIndex, name: string): Cell {
  const i = ci(idx, name);
  return i < 0 || i >= row.length ? "" : row[i];
}

// ---------------------------------------------------------------------------
// Ticket model
// ---------------------------------------------------------------------------
interface SlaState {
  due: number;
  at: number | null;
  state: string; // met | late | met-nodate | pending | due-today | breached
  bdOver: number;
  bdLeft: number;
  bdTaken: number | null;
}
interface Action {
  code: string;
  text: string;
  sev: number; // 1 critical, 2 warning, 3 info
  who: string; // owner | arietis | ops
}
interface Ticket {
  rowIndex: number;
  id: number;
  submitted: number | null;
  ownerEmail: string;   // Forms submitter (My Tickets = created by me)
  ownerName: string;
  assignee: string;     // Owner Email (digest recipient), defaults to submitter
  dept: string;
  clinic: string;
  patient: string;
  mrn: string;
  urgency: string;
  source: string;
  taskTypes: string[];
  amount: number | null;
  notes: string;
  attachments: string;
  status: string;
  firstReplyDate: number | null;
  outreachDate: number | null;
  resolutionDate: number | null;
  falseVerif: string;
  serviceRecovery: string;
  errorSource: string;
  outcome: string;
  ehr: boolean;
  link: string;
  convId: string;
  sentAt: number | null;
  firstReplyAt: number | null;
  lastReplyAt: number | null;
  replyCount: number;
  cc: string[];
  updBy: string;
  updAt: number | null;
  closedAt: number | null;
  followUpAt: number | null;
  conflicts: string;
  workNotes: string;
  opsFlag: string;
  opsReason: string;
  opsBy: string;
  opsAt: number | null;
  commitment: string;
  followUpDue: number | null;
  reviewNotes: string;
  reviewedAt: number | null;
  opsEscalated: boolean;
  opsWhy: string[];
  // derived
  isOpen: boolean;
  stage: string;
  receipt: SlaState;
  resolution: SlaState;
  lastActivity: number;
  ageBD: number;
  actions: Action[];
  flags: string[];
  topSev: number;
}

function splitList(v: Cell): string[] {
  return norm(v).split(/[;,]/).map((x) => x.trim()).filter((x) => x.length > 0);
}
function normUrgency(v: Cell): string {
  const s = norm(v).replace(/\*/g, "").trim();
  if (!s) return "Normal";
  const k = s.toLowerCase();
  if (k.indexOf("crit") === 0) return "Critical";
  if (k.indexOf("high") === 0) return "High";
  return s;
}
function toNum(v: Cell): number | null {
  if (typeof v === "number") return v;
  const s = norm(v).replace(/[$,]/g, "");
  if (!s) return null;
  const n = parseFloat(s);
  return isNaN(n) ? null : n;
}
function toBool(v: Cell): boolean {
  if (typeof v === "boolean") return v;
  const k = normKey(v);
  return k === "true" || k === "yes" || k === "1" || k === "x";
}

function rowToTicket(row: Cell[], idx: ColIndex, rowIndex: number): Ticket | null {
  const idRaw = cell(row, idx, COL.id);
  const id = toNum(idRaw);
  if (id === null) return null;
  const status = norm(cell(row, idx, COL.status)) || STATUS.submitted;
  const t: Ticket = {
    rowIndex: rowIndex,
    id: id,
    submitted: toSerial(cell(row, idx, COL.submitted)),
    ownerEmail: normKey(cell(row, idx, COL.email)),
    ownerName: norm(cell(row, idx, COL.name)),
    assignee: normKey(cell(row, idx, COL.ownerEmail)) || normKey(cell(row, idx, COL.email)),
    dept: norm(cell(row, idx, COL.dept)),
    clinic: norm(cell(row, idx, COL.clinic)) || "(No clinic)",
    patient: norm(cell(row, idx, COL.patient)),
    mrn: norm(cell(row, idx, COL.mrn)),
    urgency: normUrgency(cell(row, idx, COL.urgency)),
    source: norm(cell(row, idx, COL.source)),
    taskTypes: splitList(cell(row, idx, COL.taskType)),
    amount: toNum(cell(row, idx, COL.amount)),
    notes: norm(cell(row, idx, COL.notes)),
    attachments: norm(cell(row, idx, COL.attachments)),
    status: status,
    firstReplyDate: toSerial(cell(row, idx, COL.firstReplyDate)),
    outreachDate: toSerial(cell(row, idx, COL.outreachDate)),
    resolutionDate: toSerial(cell(row, idx, COL.resolutionDate)),
    falseVerif: norm(cell(row, idx, COL.falseVerif)),
    serviceRecovery: norm(cell(row, idx, COL.serviceRecovery)),
    errorSource: norm(cell(row, idx, COL.errorSource)),
    outcome: norm(cell(row, idx, COL.outcome)),
    ehr: toBool(cell(row, idx, COL.ehr)),
    link: norm(cell(row, idx, COL.link)),
    convId: norm(cell(row, idx, COL.convId)),
    sentAt: toSerial(cell(row, idx, COL.sentAt)),
    firstReplyAt: toSerial(cell(row, idx, COL.firstReplyAt)),
    lastReplyAt: toSerial(cell(row, idx, COL.lastReplyAt)),
    replyCount: toNum(cell(row, idx, COL.replyCount)) || 0,
    cc: splitList(cell(row, idx, COL.cc)).map((x) => x.toLowerCase()),
    updBy: norm(cell(row, idx, COL.updBy)),
    updAt: toSerial(cell(row, idx, COL.updAt)),
    closedAt: toSerial(cell(row, idx, COL.closedAt)),
    followUpAt: toSerial(cell(row, idx, COL.followUpAt)),
    conflicts: norm(cell(row, idx, COL.conflicts)),
    workNotes: norm(cell(row, idx, COL.workNotes)),
    opsFlag: norm(cell(row, idx, COL.opsFlag)),
    opsReason: norm(cell(row, idx, COL.opsReason)),
    opsBy: norm(cell(row, idx, COL.opsBy)),
    opsAt: toSerial(cell(row, idx, COL.opsAt)),
    commitment: norm(cell(row, idx, COL.commitment)),
    followUpDue: toSerial(cell(row, idx, COL.followUpDue)),
    reviewNotes: norm(cell(row, idx, COL.reviewNotes)),
    reviewedAt: toSerial(cell(row, idx, COL.reviewedAt)),
    opsEscalated: false,
    opsWhy: [],
    isOpen: true,
    stage: "",
    receipt: { due: 0, at: null, state: "pending", bdOver: 0, bdLeft: 0, bdTaken: null },
    resolution: { due: 0, at: null, state: "pending", bdOver: 0, bdLeft: 0, bdTaken: null },
    lastActivity: 0,
    ageBD: 0,
    actions: [],
    flags: [],
    topSev: 9,
  };
  return t;
}

interface EvalCtx {
  now: number;
  hol: Set<number>;
  receiptDays: number;
  resolutionDays: number;
  staleDays: number;
}

function isClosedStatus(s: string): boolean {
  const k = normKey(s);
  for (const c of CLOSED_STATUSES) if (normKey(c) === k) return true;
  return k.indexOf("closed") === 0 || k.indexOf("resolved") === 0;
}
function stageOf(status: string): string {
  const k = normKey(status);
  if (isClosedStatus(status)) return "Closed";
  if (k.indexOf("pending patient call") >= 0) return "Pending call";
  if (k.indexOf("confirmed") >= 0) return "With Arietis";
  return "Awaiting receipt";
}
const STAGES: string[] = ["Awaiting receipt", "With Arietis", "Pending call", "Closed"];

function slaFor(start: number, days: number, at: number | null, doneNoDate: boolean, ctx: EvalCtx): SlaState {
  const due = addBizDays(start, days, ctx.hol);
  const today = Math.floor(ctx.now);
  const st: SlaState = { due: due, at: at, state: "pending", bdOver: 0, bdLeft: 0, bdTaken: null };
  if (at !== null) {
    st.state = Math.floor(at) <= due ? "met" : "late";
    st.bdTaken = Math.max(0, bizDaysBetween(start, at, ctx.hol));
    if (st.state === "late") st.bdOver = bizDaysBetween(due, at, ctx.hol);
  } else if (doneNoDate) {
    st.state = "met-nodate";
  } else if (today < due) {
    st.state = "pending";
    st.bdLeft = bizDaysBetween(today, due, ctx.hol);
  } else if (today === due) {
    st.state = "due-today";
  } else {
    st.state = "breached";
    st.bdOver = Math.max(1, bizDaysBetween(due, today, ctx.hol));
  }
  return st;
}

function evaluateTicket(t: Ticket, ctx: EvalCtx): Ticket {
  const start = t.submitted !== null ? t.submitted : ctx.now;
  t.isOpen = !isClosedStatus(t.status);
  t.stage = stageOf(t.status);
  const why: string[] = [];
  const fk = normKey(t.opsFlag);
  if (fk === "yes") why.push(t.opsReason ? "Escalated: " + t.opsReason : "Escalated by ops");
  if (t.urgency === "Critical") why.push("Critical");
  if (normKey(t.serviceRecovery) === "yes") why.push("Service recovery");
  t.opsEscalated = fk !== "cleared" && why.length > 0;
  t.opsWhy = t.opsEscalated ? why : [];

  const rcptAt = t.firstReplyAt !== null ? t.firstReplyAt : t.firstReplyDate;
  const rcptByStatus = t.stage !== "Awaiting receipt";
  t.receipt = slaFor(start, ctx.receiptDays, rcptAt, rcptByStatus, ctx);

  let resAt: number | null = null;
  if (!t.isOpen) resAt = t.resolutionDate !== null ? t.resolutionDate : t.closedAt !== null ? t.closedAt : t.updAt;
  t.resolution = slaFor(start, ctx.resolutionDays, resAt, !t.isOpen, ctx);

  t.lastActivity = Math.max(start, t.updAt || 0, t.lastReplyAt || 0, t.sentAt || 0, t.followUpAt || 0, t.outreachDate || 0);
  t.ageBD = Math.max(0, bizDaysBetween(start, t.isOpen ? ctx.now : (resAt !== null ? resAt : ctx.now), ctx.hol));

  const a: Action[] = [];
  if (t.isOpen) {
    if (t.receipt.state === "breached") {
      a.push({ code: "RCPT_OVERDUE", text: "No receipt from Arietis — " + t.receipt.bdOver + " BD past SLA", sev: 1, who: "arietis" });
    }
    if (t.resolution.state === "breached") {
      a.push({ code: "RES_OVERDUE", text: "Resolution " + t.resolution.bdOver + " BD past SLA — escalate", sev: 1, who: "arietis" });
    }
    if (t.receipt.state === "due-today") a.push({ code: "RCPT_DUE", text: "Arietis receipt due today", sev: 2, who: "arietis" });
    if (t.resolution.state === "due-today") a.push({ code: "RES_DUE", text: "Resolution due today", sev: 2, who: "arietis" });
    if (t.stage === "Awaiting receipt" && rcptAt !== null) {
      a.push({ code: "REPLY_UNLOGGED", text: "Arietis replied " + fmtDateTime(rcptAt) + " — update status", sev: 2, who: "owner" });
    } else if (t.lastReplyAt !== null && (t.updAt === null || t.lastReplyAt > t.updAt + 1 / 1440)) {
      a.push({ code: "NEW_REPLY", text: "New Arietis reply " + fmtDateTime(t.lastReplyAt) + " — review & update", sev: 2, who: "owner" });
    }
    if (t.clinic === "(No clinic)" || (!t.patient && !t.mrn)) {
      a.push({ code: "MISSING_DETAILS", text: "Clinic/patient blank (likely wiped by a Form edit) — fill in", sev: 2, who: "owner" });
    }
    if (t.conflicts) {
      a.push({ code: "SYNC_CONFLICT", text: "Review form/Master conflict: " + t.conflicts, sev: 2, who: "owner" });
    }
    if (t.stage === "Pending call" && t.outreachDate === null) {
      a.push({ code: "OUTREACH", text: "Log date of patient outreach", sev: 2, who: "owner" });
    }
    if (t.receipt.state === "met-nodate") {
      a.push({ code: "RCPT_DATE", text: "Add date of first Arietis reply", sev: 3, who: "owner" });
    }
    if (t.opsEscalated && t.followUpDue !== null) {
      const fd = Math.floor(t.followUpDue);
      const today = Math.floor(ctx.now);
      if (fd < today) a.push({ code: "OPS_FU_OVERDUE", text: "Arietis follow-up overdue since " + fmtShort(fd) + (t.commitment ? " — " + t.commitment : ""), sev: 1, who: "arietis" });
      else if (fd === today) a.push({ code: "OPS_FU_DUE", text: "Arietis follow-up due today" + (t.commitment ? " — " + t.commitment : ""), sev: 2, who: "arietis" });
    }
    if (t.opsEscalated && t.reviewedAt === null) a.push({ code: "OPS_REVIEW", text: "Ops escalation — not yet reviewed with Arietis", sev: 3, who: "ops" });
    const idle = bizDaysBetween(t.lastActivity, ctx.now, ctx.hol);
    const hasBreach = a.some((x) => x.sev === 1);
    if (!hasBreach && idle >= ctx.staleDays) {
      a.push({ code: "STALE", text: "No activity for " + idle + " BD", sev: 3, who: "owner" });
    }
  } else {
    const missing: string[] = [];
    if (t.resolutionDate === null) missing.push("resolution date");
    if (!t.errorSource) missing.push("source of error");
    if (!t.outcome && normKey(t.status) !== normKey(STATUS.closedNoAction)) missing.push("outcome");
    if (!t.falseVerif) missing.push("false verification Y/N");
    if (missing.length) {
      a.push({ code: "CLOSE_INCOMPLETE", text: "Closed — missing " + missing.join(", "), sev: 3, who: "owner" });
    }
  }
  a.sort((x, y) => x.sev - y.sev);
  t.actions = a;
  t.topSev = a.length ? a[0].sev : 9;

  const f: string[] = [];
  if (t.urgency === "Critical") f.push("Critical");
  else if (t.urgency === "High") f.push("High");
  if (normKey(t.serviceRecovery) === "yes") f.push("Service recovery");
  if (normKey(t.falseVerif) === "yes") f.push("False verification");
  if (t.opsEscalated && t.isOpen) f.push("Ops escalation");
  if (t.ehr) f.push("EHR task");
  t.flags = f;
  return t;
}

function holidaySet(serials: Cell[]): Set<number> {
  const s = new Set<number>();
  for (const v of serials) {
    const n = toSerial(v);
    if (n !== null) s.add(Math.floor(n));
  }
  return s;
}

interface Settings {
  receiptDays: number;
  resolutionDays: number;
  staleDays: number;
  arietisEmail: string;
  billingInbox: string;
  leadership: string[];
  clinicDigest: boolean;
  arietisDigest: boolean;
  workbookUrl: string;
  formUrl: string;
  digestIds: boolean;
}
function readSettings(rows: Cell[][]): Settings {
  const kv = new Map<string, string>();
  for (const r of rows) if (r.length >= 2 && !isBlank(r[0])) kv.set(normKey(r[0]).replace(/[^a-z0-9]/g, ""), norm(r[1]));
  const num = (k: string, d: number): number => {
    const v = parseFloat(kv.get(k) || "");
    return isNaN(v) ? d : v;
  };
  const yes = (k: string): boolean => normKey(kv.get(k) || "") === "yes";
  return {
    receiptDays: num("receiptslabusinessdays", DEFAULTS.receiptDays),
    resolutionDays: num("resolutionslabusinessdays", DEFAULTS.resolutionDays),
    staleDays: num("staleafterbusinessdays", DEFAULTS.staleDays),
    arietisEmail: kv.get("arietisemail") || DEFAULTS.arietisEmail,
    billingInbox: kv.get("billinginbox") || DEFAULTS.billingInbox,
    leadership: splitList(kv.get("leadershiprecipients") || ""),
    clinicDigest: yes("clinicdailydigest"),
    arietisDigest: yes("arietisdailyfollowup"),
    workbookUrl: kv.get("workbookurl") || "",
    formUrl: kv.get("formurl") || "",
    digestIds: yes("digestincludepatientids"),
  };
}

function loadTickets(headers: Cell[], rows: Cell[][], ctx: EvalCtx): Ticket[] {
  const idx = buildIndex(headers);
  const out: Ticket[] = [];
  for (let i = 0; i < rows.length; i++) {
    const t = rowToTicket(rows[i], idx, i);
    if (t) out.push(evaluateTicket(t, ctx));
  }
  return out;
}

// ---------------------------------------------------------------------------
// Stats
// ---------------------------------------------------------------------------
interface Stats {
  key: string;
  total: number;
  open: number;
  closed: number;
  stage: number[]; // counts per STAGES
  rcptOnTime: number;
  rcptMiss: number;
  resOnTime: number;
  resMiss: number;
  rcptBdSum: number;
  rcptBdN: number;
  resBdSum: number;
  resBdN: number;
  breachedOpen: number;
  amountOpen: number;
  falseVerif: number;
  serviceRecovery: number;
  needsAction: number;
}
function emptyStats(key: string): Stats {
  return {
    key: key, total: 0, open: 0, closed: 0, stage: [0, 0, 0, 0], rcptOnTime: 0, rcptMiss: 0, resOnTime: 0,
    resMiss: 0, rcptBdSum: 0, rcptBdN: 0, resBdSum: 0, resBdN: 0, breachedOpen: 0, amountOpen: 0,
    falseVerif: 0, serviceRecovery: 0, needsAction: 0,
  };
}
function addToStats(s: Stats, t: Ticket): void {
  s.total += 1;
  if (t.isOpen) s.open += 1;
  else s.closed += 1;
  const si = STAGES.indexOf(t.stage);
  if (si >= 0) s.stage[si] += 1;
  const r = t.receipt.state;
  if (r === "met") s.rcptOnTime += 1;
  if (r === "late" || r === "breached") s.rcptMiss += 1;
  const x = t.resolution.state;
  if (x === "met") s.resOnTime += 1;
  if (x === "late" || x === "breached") s.resMiss += 1;
  if (t.receipt.bdTaken !== null) { s.rcptBdSum += t.receipt.bdTaken; s.rcptBdN += 1; }
  if (t.resolution.bdTaken !== null) { s.resBdSum += t.resolution.bdTaken; s.resBdN += 1; }
  if (t.isOpen && (r === "breached" || x === "breached")) s.breachedOpen += 1;
  if (t.isOpen && t.amount !== null) s.amountOpen += t.amount;
  if (normKey(t.falseVerif) === "yes") s.falseVerif += 1;
  if (normKey(t.serviceRecovery) === "yes") s.serviceRecovery += 1;
  if (t.actions.length) s.needsAction += 1;
}
function pct(a: number, b: number): number | null {
  return b > 0 ? Math.round((a / b) * 100) : null;
}
function computeStats(tickets: Ticket[]): { total: Stats; byClinic: Stats[] } {
  const total = emptyStats("All clinics");
  const m = new Map<string, Stats>();
  for (const t of tickets) {
    addToStats(total, t);
    let s = m.get(t.clinic);
    if (!s) { s = emptyStats(t.clinic); m.set(t.clinic, s); }
    addToStats(s, t);
  }
  const arr: Stats[] = [];
  m.forEach((v) => arr.push(v));
  arr.sort((a, b) => b.open - a.open || b.total - a.total || a.key.localeCompare(b.key));
  return { total: total, byClinic: arr };
}

// ---------------------------------------------------------------------------
// Raw Data -> Master sync
// ---------------------------------------------------------------------------
interface CellChange {
  col: string;
  colIndex: number;
  oldVal: Cell;
  newVal: Cell;
  logged: boolean;
}
interface RowUpdate {
  rowIndex: number;
  id: number;
  by: string;
  changes: CellChange[];
}
interface SyncPlan {
  appends: Cell[][];
  appendIds: number[];
  updates: RowUpdate[];
  baselined: number;
  duplicateIds: number[];
}
function cellEq(a: Cell, b: Cell): boolean {
  if (isBlank(a) && isBlank(b)) return true;
  if (typeof a === "number" && typeof b === "number") return Math.abs(a - b) < 1e-6;
  const na = toNum(a);
  const nb = toNum(b);
  if (typeof a === "number" || typeof b === "number") {
    if (na !== null && nb !== null) return Math.abs(na - nb) < 1e-6;
  }
  return norm(a) === norm(b);
}

/**
 * Plans the Raw Data -> Master sync. Rules:
 *  - New Ids in Raw Data are appended to Master (all Form columns + EHR Task + response link).
 *  - Existing Ids with RawSyncedAt blank are BASELINED: RawSyncedAt is stamped, nothing else
 *    touched (Master was curated by hand before the add-in existed).
 *  - Existing Ids whose Raw "Completion time" is newer than RawSyncedAt were edited through the
 *    Form's edit link: only NON-BLANK mergeable fields are copied; nothing is ever blanked;
 *    submission time / owner / Id are never overwritten.
 */
function planSync(
  rawHeaders: Cell[], rawRows: Cell[][], rawLinks: Map<number, string>,
  masterHeaders: Cell[], masterRows: Cell[][]
): SyncPlan {
  const ri = buildIndex(rawHeaders);
  const mi = buildIndex(masterHeaders);
  const plan: SyncPlan = { appends: [], appendIds: [], updates: [], baselined: 0, duplicateIds: [] };
  const masterById = new Map<number, number>();
  for (let i = 0; i < masterRows.length; i++) {
    const id = toNum(cell(masterRows[i], mi, COL.id));
    if (id === null) continue;
    if (masterById.has(id)) plan.duplicateIds.push(id);
    else masterById.set(id, i);
  }
  const width = masterHeaders.length;
  for (const r of rawRows) {
    const id = toNum(cell(r, ri, COL.id));
    if (id === null) continue;
    const rawDone = toSerial(cell(r, ri, COL.submitted));
    const rawLink = rawLinks.get(id) || "";
    const mIdx = masterById.get(id);
    if (mIdx === undefined) {
      const row: Cell[] = [];
      for (let c = 0; c < width; c++) row.push("");
      for (const name of FORM_COLS) {
        const tc = ci(mi, name);
        if (tc >= 0) row[tc] = cell(r, ri, name);
      }
      const set = (name: string, v: Cell): void => {
        const tc = ci(mi, name);
        if (tc >= 0) row[tc] = v;
      };
      set(COL.ehr, toBool(cell(r, ri, COL.ehr)));
      set(COL.link, rawLink);
      set(COL.replyCount, 0);
      set(COL.updBy, normKey(cell(r, ri, COL.email)) + " (form)");
      set(COL.updAt, rawDone !== null ? rawDone : "");
      set(COL.rawSynced, rawDone !== null ? rawDone : "");
      if (isBlank(cell(r, ri, COL.status))) set(COL.status, STATUS.submitted);
      plan.appends.push(row);
      plan.appendIds.push(id);
      masterById.set(id, -1);
      continue;
    }
    if (mIdx < 0) continue; // duplicate raw id already appended in this pass
    const mrow = masterRows[mIdx];
    const synced = toSerial(cell(mrow, mi, COL.rawSynced));
    const changes: CellChange[] = [];
    const push = (name: string, v: Cell, logged: boolean): void => {
      const tc = ci(mi, name);
      if (tc < 0) return;
      const old = mrow[tc];
      if (!cellEq(old, v)) changes.push({ col: name, colIndex: tc, oldVal: old, newVal: v, logged: logged });
    };
    if (synced === null) {
      if (rawDone !== null) push(COL.rawSynced, rawDone, false);
      if (rawLink && isBlank(cell(mrow, mi, COL.link))) push(COL.link, rawLink, false);
      // EHR Task was tracked only in Raw Data before setup; carry it over once.
      if (isBlank(cell(mrow, mi, COL.ehr))) push(COL.ehr, toBool(cell(r, ri, COL.ehr)), false);
      if (isBlank(cell(mrow, mi, COL.replyCount))) push(COL.replyCount, 0, false);
      if (changes.length) plan.updates.push({ rowIndex: mIdx, id: id, by: "setup", changes: changes });
      plan.baselined += 1;
      continue;
    }
    if (rawDone === null || rawDone <= synced + 1e-6) continue;
    for (const name of MERGEABLE_COLS) {
      const v = cell(r, ri, name);
      if (isBlank(v)) continue;
      push(name, v, true);
    }
    if (toBool(cell(r, ri, COL.ehr)) && !toBool(cell(mrow, mi, COL.ehr))) push(COL.ehr, true, true);
    if (rawLink && isBlank(cell(mrow, mi, COL.link))) push(COL.link, rawLink, false);
    const editor = normKey(cell(r, ri, COL.email));
    const hasLogged = changes.some((c) => c.logged);
    if (hasLogged) {
      push(COL.updBy, editor + " (form edit)", false);
      push(COL.updAt, rawDone, false);
    }
    push(COL.rawSynced, rawDone, false);
    if (changes.length) plan.updates.push({ rowIndex: mIdx, id: id, by: editor + " (form edit)", changes: changes });
  }
  return plan;
}

// ---------------------------------------------------------------------------
// Arietis reply logging (used by the inbox flow via Office Script)
// ---------------------------------------------------------------------------
interface ReplyResult {
  matched: boolean;
  logged: boolean;
  ticketId: number;
  reason: string;
  ownerEmail: string;
  firstReply: boolean;
  changes: CellChange[];
  rowIndex: number;
}
function ticketIdFromSubject(subject: string): number | null {
  const m = subject.match(/\[BE-(\d+)\]/i) || subject.match(/Escalation\s*#\s*(\d+)/i) || subject.match(/Ticket\s*#\s*(\d+)/i);
  return m ? parseInt(m[1], 10) : null;
}
function planReply(
  masterHeaders: Cell[], masterRows: Cell[][], conversationId: string, receivedAt: number, subject: string
): ReplyResult {
  const mi = buildIndex(masterHeaders);
  const res: ReplyResult = { matched: false, logged: false, ticketId: 0, reason: "", ownerEmail: "", firstReply: false, changes: [], rowIndex: -1 };
  const conv = norm(conversationId);
  let hit = -1;
  if (conv) {
    for (let i = 0; i < masterRows.length; i++) {
      if (norm(cell(masterRows[i], mi, COL.convId)) === conv) { hit = i; break; }
    }
  }
  if (hit < 0) {
    const sid = ticketIdFromSubject(subject || "");
    if (sid !== null) {
      for (let i = 0; i < masterRows.length; i++) {
        if (toNum(cell(masterRows[i], mi, COL.id)) === sid) { hit = i; break; }
      }
    }
  }
  if (hit < 0) { res.reason = "no matching ticket"; return res; }
  const row = masterRows[hit];
  res.matched = true;
  res.rowIndex = hit;
  res.ticketId = toNum(cell(row, mi, COL.id)) || 0;
  res.ownerEmail = normKey(cell(row, mi, COL.email));
  if (isClosedStatus(norm(cell(row, mi, COL.status)))) { res.reason = "ticket closed — not logged"; return res; }
  const push = (name: string, v: Cell, logged: boolean): void => {
    const tc = ci(mi, name);
    if (tc < 0) return;
    if (!cellEq(row[tc], v)) res.changes.push({ col: name, colIndex: tc, oldVal: row[tc], newVal: v, logged: logged });
  };
  const first = toSerial(cell(row, mi, COL.firstReplyAt));
  const last = toSerial(cell(row, mi, COL.lastReplyAt));
  if (first === null || receivedAt < first) {
    push(COL.firstReplyAt, receivedAt, true);
    res.firstReply = first === null;
    if (isBlank(cell(row, mi, COL.firstReplyDate))) push(COL.firstReplyDate, Math.floor(receivedAt), true);
  }
  if (last === null || receivedAt > last) push(COL.lastReplyAt, receivedAt, true);
  if (conv && isBlank(cell(row, mi, COL.convId))) push(COL.convId, conv, false);
  const cnt = toNum(cell(row, mi, COL.replyCount)) || 0;
  push(COL.replyCount, cnt + 1, false);
  res.logged = true;
  return res;
}

// ---------------------------------------------------------------------------
// Email builders (shared: add-in reminders + Power Automate digests)
// ---------------------------------------------------------------------------
interface Email {
  to: string;
  cc: string;
  subject: string;
  html: string;
  kind: string;
}
interface ActivityRow {
  at: number;
  id: number;
  field: string;
  oldVal: string;
  newVal: string;
  by: string;
  source: string;
}

const EM = {
  ink: "#1b1f24", muted: "#5b6470", line: "#d9dee4", head: "#f3f5f7",
  red: "#b42318", amber: "#a15c07", green: "#067647", blue: "#1f5f8b",
};
function slaText(s: SlaState): string {
  switch (s.state) {
    case "met": return "Met " + fmtShort(s.at);
    case "late": return "Late " + fmtShort(s.at) + " (+" + s.bdOver + " BD)";
    case "met-nodate": return "Met (no date)";
    case "due-today": return "Due today";
    case "breached": return "Overdue " + s.bdOver + " BD";
    default: return "Due " + weekdayName(s.due) + " " + fmtShort(s.due);
  }
}
function slaColor(s: SlaState): string {
  if (s.state === "breached" || s.state === "late") return EM.red;
  if (s.state === "due-today") return EM.amber;
  if (s.state === "met" || s.state === "met-nodate") return EM.green;
  return EM.muted;
}
function ticketLabel(t: Ticket): string {
  return "#" + t.id;
}
/** Owner/leadership digests omit patient code + MRN by default (repo privacy stance); clinic recaps include them. */
let SHOW_IDS = false;
function patientLabel(t: Ticket): string {
  if (!SHOW_IDS) return t.taskTypes.length ? t.taskTypes[0].replace(/\s*\(.*\)$/, "") : "Billing escalation";
  return [t.patient, t.mrn ? "MRN " + t.mrn : ""].filter((x) => x).join(" · ");
}
function pHdr(): string { return SHOW_IDS ? "Patient" : "Issue"; }
function hTable(headers: string[], rows: string[][]): string {
  if (!rows.length) return "";
  let h = '<table cellpadding="0" cellspacing="0" style="border-collapse:collapse;width:100%;font-size:13px;margin:6px 0 14px">';
  h += "<tr>" + headers.map((x) => '<th align="left" style="background:' + EM.head + ";color:" + EM.muted + ";font-weight:600;padding:6px 8px;border-bottom:1px solid " + EM.line + ';font-size:12px">' + escHtml(x) + "</th>").join("") + "</tr>";
  for (const r of rows) {
    h += "<tr>" + r.map((c) => '<td style="padding:6px 8px;border-bottom:1px solid ' + EM.line + ';vertical-align:top">' + c + "</td>").join("") + "</tr>";
  }
  return h + "</table>";
}
function hSection(title: string, body: string, count: number | null): string {
  if (!body) return "";
  const c = count === null ? "" : ' <span style="color:' + EM.muted + ';font-weight:400">(' + count + ")</span>";
  return '<h3 style="font-size:14px;margin:18px 0 4px;color:' + EM.ink + '">' + escHtml(title) + c + "</h3>" + body;
}
function hKpis(items: { label: string; value: string; color: string }[]): string {
  let h = '<table cellpadding="0" cellspacing="0" style="border-collapse:collapse;margin:8px 0 6px"><tr>';
  for (const k of items) {
    h += '<td style="padding:8px 14px 8px 0"><div style="font-size:22px;font-weight:700;color:' + k.color + '">' + escHtml(k.value) + '</div><div style="font-size:11px;color:' + EM.muted + ';text-transform:uppercase;letter-spacing:.04em">' + escHtml(k.label) + "</div></td>";
  }
  return h + "</tr></table>";
}
function hWrap(title: string, sub: string, body: string, s: Settings): string {
  const link = s.workbookUrl
    ? '<p style="margin:18px 0 0"><a href="' + escHtml(s.workbookUrl) + '" style="background:' + EM.blue + ';color:#fff;text-decoration:none;padding:8px 14px;border-radius:4px;font-size:13px;font-weight:600">Open escalation tracker</a></p>'
    : "";
  return '<div style="font-family:Segoe UI,Arial,sans-serif;color:' + EM.ink + ';max-width:760px">' +
    '<div style="font-size:11px;color:' + EM.muted + ';text-transform:uppercase;letter-spacing:.06em">Patient Billing Escalations</div>' +
    '<h2 style="font-size:18px;margin:2px 0 2px">' + escHtml(title) + "</h2>" +
    '<div style="font-size:13px;color:' + EM.muted + '">' + escHtml(sub) + "</div>" + body + link +
    '<p style="font-size:11px;color:' + EM.muted + ';margin-top:20px">SLA: Arietis confirms receipt within ' + s.receiptDays + " business days and resolves within " + s.resolutionDays + " business days of submission. Automated message from the billing escalation tracker.</p></div>";
}
function sevDot(sev: number): string {
  const c = sev === 1 ? EM.red : sev === 2 ? EM.amber : EM.muted;
  return '<span style="color:' + c + ';font-weight:700">●</span> ';
}
function actionRows(ts: Ticket[]): string[][] {
  return ts.filter((t) => t.actions.length > 0).map((t) => [
    "<b>" + ticketLabel(t) + "</b>",
    escHtml(t.clinic),
    escHtml(patientLabel(t)),
    t.actions.map((a) => sevDot(a.sev) + escHtml(a.text)).join("<br>"),
  ]);
}
function openRows(ts: Ticket[]): string[][] {
  return ts.map((t) => [
    "<b>" + ticketLabel(t) + "</b>" + (t.flags.length ? '<div style="font-size:11px;color:' + EM.red + '">' + escHtml(t.flags.filter((f) => f !== "EHR task").join(" · ")) + "</div>" : ""),
    escHtml(t.clinic),
    escHtml(patientLabel(t)),
    escHtml(t.status),
    '<span style="color:' + slaColor(t.receipt) + '">' + escHtml(slaText(t.receipt)) + "</span>",
    '<span style="color:' + slaColor(t.resolution) + '">' + escHtml(slaText(t.resolution)) + "</span>",
  ]);
}
function sortForAttention(ts: Ticket[]): Ticket[] {
  const urg = (u: string): number => (u === "Critical" ? 0 : u === "High" ? 1 : 2);
  return ts.slice().sort((a, b) => a.topSev - b.topSev || urg(a.urgency) - urg(b.urgency) || (a.submitted || 0) - (b.submitted || 0));
}
function involves(t: Ticket, email: string): boolean {
  return t.assignee === email || t.cc.indexOf(email) >= 0;
}

/** Builds digest emails. mode: "open" | "close" | "weekly". */
function buildDigests(tickets: Ticket[], activity: ActivityRow[], ctx: EvalCtx, s: Settings, mode: string): Email[] {
  SHOW_IDS = s.digestIds;
  const out: Email[] = [];
  const today = Math.floor(ctx.now);
  const dayLabel = weekdayName(today) + " " + fmtDate(today);
  const people = new Map<string, string>();
  for (const t of tickets) {
    if (t.assignee && !people.has(t.assignee)) people.set(t.assignee, t.assignee === t.ownerEmail ? t.ownerName : t.assignee);
    for (const c of t.cc) if (c.indexOf("@") > 0 && !people.has(c)) people.set(c, c);
  }
  const weekStart = today - 6;

  people.forEach((name, email) => {
    const mine = tickets.filter((t) => involves(t, email));
    const open = sortForAttention(mine.filter((t) => t.isOpen));
    const needs = sortForAttention(mine.filter((t) => t.actions.length > 0));
    const first = (name || email).split(" ")[0];
    const breached = open.filter((t) => t.receipt.state === "breached" || t.resolution.state === "breached");
    const dueToday = open.filter((t) => t.receipt.state === "due-today" || t.resolution.state === "due-today");
    let body = "";
    let subject = "";
    if (mode === "open") {
      if (!open.length && !needs.length) return;
      subject = "Billing escalations — " + open.length + " open" + (needs.length ? ", " + needs.length + " need action" : "") + " (" + fmtShort(today) + ")";
      body += hKpis([
        { label: "Open", value: String(open.length), color: EM.ink },
        { label: "Need action", value: String(needs.length), color: needs.length ? EM.amber : EM.ink },
        { label: "Past SLA", value: String(breached.length), color: breached.length ? EM.red : EM.ink },
        { label: "Due today", value: String(dueToday.length), color: dueToday.length ? EM.amber : EM.ink },
      ]);
      body += hSection("Needs your action", hTable(["Ticket", "Clinic", pHdr(), "Action"], actionRows(needs)), needs.length);
      body += hSection("Your open tickets", hTable(["Ticket", "Clinic", pHdr(), "Status", "Receipt (2 BD)", "Resolution (5 BD)"], openRows(open)), open.length);
      out.push({ to: email, cc: "", subject: subject, html: hWrap("Good morning, " + first, dayLabel + " · start-of-day status", body, s), kind: "owner-open" });
    } else if (mode === "close") {
      const ids = new Set<number>(mine.map((t) => t.id));
      const todays = activity.filter((a) => Math.floor(a.at) === today && ids.has(a.id));
      const closedToday = mine.filter((t) => !t.isOpen && t.resolution.at !== null && Math.floor(t.resolution.at) === today);
      const repliedToday = mine.filter((t) => t.lastReplyAt !== null && Math.floor(t.lastReplyAt) === today);
      const nextBiz = addBizDays(today, 1, ctx.hol);
      const dueTomorrow = open.filter((t) => (t.receipt.state === "pending" && t.receipt.due === nextBiz) || (t.resolution.state === "pending" && t.resolution.due === nextBiz));
      if (!open.length && !todays.length && !closedToday.length) return;
      subject = "Billing escalations — end of day " + fmtShort(today) + ": " + open.length + " open" + (breached.length ? ", " + breached.length + " past SLA" : "");
      body += hKpis([
        { label: "Arietis replies today", value: String(repliedToday.length), color: EM.ink },
        { label: "Closed today", value: String(closedToday.length), color: closedToday.length ? EM.green : EM.ink },
        { label: "Still open", value: String(open.length), color: EM.ink },
        { label: "Past SLA", value: String(breached.length), color: breached.length ? EM.red : EM.ink },
      ]);
      body += hSection("Arietis replied today", hTable(["Ticket", "Clinic", pHdr(), "Last reply", "Status"], repliedToday.map((t) => [
        "<b>" + ticketLabel(t) + "</b>", escHtml(t.clinic), escHtml(patientLabel(t)), escHtml(fmtTime(t.lastReplyAt)), escHtml(t.status),
      ])), repliedToday.length);
      body += hSection("Changes today", hTable(["Ticket", "Field", "New value", "By"], todays.slice(0, 40).map((a) => [
        "<b>#" + a.id + "</b>", escHtml(a.field), escHtml(a.newVal), escHtml(a.by),
      ])), todays.length);
      body += hSection("Still needs action", hTable(["Ticket", "Clinic", pHdr(), "Action"], actionRows(needs)), needs.length);
      body += hSection("SLA due next business day (" + fmtShort(nextBiz) + ")", hTable(["Ticket", "Clinic", pHdr(), "Status", "Receipt", "Resolution"], openRows(dueTomorrow)), dueTomorrow.length);
      out.push({ to: email, cc: "", subject: subject, html: hWrap("End of day, " + first, dayLabel + " · close-of-business status", body, s), kind: "owner-close" });
    } else if (mode === "weekly") {
      const opened = mine.filter((t) => t.submitted !== null && Math.floor(t.submitted) >= weekStart);
      const closed = mine.filter((t) => !t.isOpen && t.resolution.at !== null && Math.floor(t.resolution.at) >= weekStart);
      if (!open.length && !opened.length && !closed.length) return;
      const st = computeStats(mine).total;
      const flagged = open.filter((t) => t.flags.some((f) => f !== "EHR task") || t.topSev === 1);
      const aging = [0, 0, 0, 0];
      for (const t of open) aging[t.ageBD <= 2 ? 0 : t.ageBD <= 5 ? 1 : t.ageBD <= 10 ? 2 : 3] += 1;
      subject = "Billing escalations — weekly recap (" + fmtShort(weekStart) + "–" + fmtShort(today) + ")";
      body += hKpis([
        { label: "Opened", value: String(opened.length), color: EM.ink },
        { label: "Closed", value: String(closed.length), color: EM.green },
        { label: "Still open", value: String(open.length), color: EM.ink },
        { label: "Receipt on time", value: pct(st.rcptOnTime, st.rcptOnTime + st.rcptMiss) === null ? "—" : pct(st.rcptOnTime, st.rcptOnTime + st.rcptMiss) + "%", color: EM.blue },
        { label: "Resolved on time", value: pct(st.resOnTime, st.resOnTime + st.resMiss) === null ? "—" : pct(st.resOnTime, st.resOnTime + st.resMiss) + "%", color: EM.blue },
      ]);
      body += hSection("Open ticket age (business days)", hTable(["0–2", "3–5", "6–10", "11+"], [aging.map((n) => String(n))]), null);
      body += hSection("Escalation flags", hTable(["Ticket", "Clinic", pHdr(), "Status", "Receipt", "Resolution"], openRows(flagged)), flagged.length);
      body += hSection("Actions needed", hTable(["Ticket", "Clinic", pHdr(), "Action"], actionRows(needs)), needs.length);
      body += hSection("Closed this week", hTable(["Ticket", "Clinic", pHdr(), "Outcome", "Resolved"], closed.map((t) => [
        "<b>" + ticketLabel(t) + "</b>", escHtml(t.clinic), escHtml(patientLabel(t)), escHtml(t.outcome || t.status), '<span style="color:' + slaColor(t.resolution) + '">' + escHtml(slaText(t.resolution)) + "</span>",
      ])), closed.length);
      out.push({ to: email, cc: "", subject: subject, html: hWrap("Weekly recap, " + first, fmtDate(weekStart) + " – " + fmtDate(today), body, s), kind: "owner-weekly" });
    }
  });

  // ---- Leadership rollup (weekly, and daily-open when something is past SLA) ----
  if (s.leadership.length && (mode === "weekly" || mode === "open")) {
    const open = sortForAttention(tickets.filter((t) => t.isOpen));
    const breached = open.filter((t) => t.topSev === 1);
    if (mode === "weekly" || breached.length) {
      const st = computeStats(tickets);
      const rows = st.byClinic.filter((c) => c.open > 0 || mode === "weekly").map((c) => [
        "<b>" + escHtml(c.key) + "</b>", String(c.open), String(c.stage[0]), String(c.stage[1] + c.stage[2]),
        '<span style="color:' + (c.breachedOpen ? EM.red : EM.ink) + '">' + c.breachedOpen + "</span>",
        (pct(c.rcptOnTime, c.rcptOnTime + c.rcptMiss) === null ? "—" : pct(c.rcptOnTime, c.rcptOnTime + c.rcptMiss) + "%"),
        (pct(c.resOnTime, c.resOnTime + c.resMiss) === null ? "—" : pct(c.resOnTime, c.resOnTime + c.resMiss) + "%"),
      ]);
      let body = hKpis([
        { label: "Open", value: String(st.total.open), color: EM.ink },
        { label: "Past SLA", value: String(st.total.breachedOpen), color: st.total.breachedOpen ? EM.red : EM.ink },
        { label: "Receipt on time", value: (pct(st.total.rcptOnTime, st.total.rcptOnTime + st.total.rcptMiss) || 0) + "%", color: EM.blue },
        { label: "Resolved on time", value: (pct(st.total.resOnTime, st.total.resOnTime + st.total.resMiss) || 0) + "%", color: EM.blue },
        { label: "$ open", value: fmtMoney(st.total.amountOpen), color: EM.ink },
      ]);
      body += hSection("By clinic", hTable(["Clinic", "Open", "Awaiting receipt", "With Arietis", "Past SLA", "Receipt on time", "Resolved on time"], rows), null);
      body += hSection("Past SLA", hTable(["Ticket", "Clinic", pHdr(), "Status", "Receipt", "Resolution"], openRows(breached)), breached.length);
      const ops = opsQueue(tickets);
      body += hSection("Ops escalations", hTable(["Ticket", "Clinic", pHdr(), "Why", "Arietis commitment", "Follow-up"], ops.map((t) => [
        "<b>" + ticketLabel(t) + "</b>", escHtml(t.clinic), escHtml(patientLabel(t)), escHtml(t.opsWhy.join("; ")), escHtml(t.commitment || "—"),
        t.followUpDue === null ? "—" : '<span style="color:' + (Math.floor(t.followUpDue) < today ? EM.red : Math.floor(t.followUpDue) === today ? EM.amber : EM.ink) + '">' + escHtml(fmtShort(t.followUpDue)) + "</span>",
      ])), ops.length);
      out.push({
        to: s.leadership.join(";"), cc: "",
        subject: mode === "weekly" ? "Billing escalations — weekly rollup (" + fmtShort(weekStart) + "–" + fmtShort(today) + ")" : "Billing escalations — " + breached.length + " past SLA (" + fmtShort(today) + ")",
        html: hWrap(mode === "weekly" ? "Weekly rollup" : "Escalation flags", dayLabel, body, s), kind: "leadership-" + mode,
      });
    }
  }
  return out;
}

/** Clinic inbox digest (open-of-business), when Settings ClinicDailyDigest = Yes. */
function buildClinicDigests(tickets: Ticket[], clinicEmails: Map<string, string>, ctx: EvalCtx, s: Settings): Email[] {
  SHOW_IDS = true;
  const out: Email[] = [];
  const by = new Map<string, Ticket[]>();
  for (const t of tickets) if (t.isOpen) {
    const a = by.get(t.clinic) || [];
    a.push(t);
    by.set(t.clinic, a);
  }
  by.forEach((ts, clinic) => {
    const to = clinicEmails.get(normKey(clinic));
    if (!to) return;
    const sorted = sortForAttention(ts);
    const body = hSection("Open escalations", hTable(["Ticket", "Clinic", pHdr(), "Status", "Receipt", "Resolution"], openRows(sorted)), sorted.length);
    out.push({ to: to, cc: "", subject: clinic + " — " + ts.length + " open billing escalations (" + fmtShort(ctx.now) + ")", html: hWrap(clinic, "Open billing escalations", body, s), kind: "clinic" });
  });
  return out;
}

/** Follow-up to Arietis listing tickets past SLA (add-in button, or daily when enabled). */
function buildArietisFollowup(ts: Ticket[], s: Settings, ctx: EvalCtx, from: string): Email | null {
  const list = sortForAttention(ts.filter((t) => t.isOpen && (t.receipt.state === "breached" || t.resolution.state === "breached" || t.receipt.state === "due-today" || t.resolution.state === "due-today")));
  if (!list.length) return null;
  const rows = list.map((t) => [
    "<b>" + ticketLabel(t) + "</b>", escHtml(t.patient), escHtml(t.mrn), escHtml(fmtDate(t.submitted)), escHtml(t.status),
    '<span style="color:' + slaColor(t.receipt) + '">' + escHtml(slaText(t.receipt)) + "</span>",
    '<span style="color:' + slaColor(t.resolution) + '">' + escHtml(slaText(t.resolution)) + "</span>",
  ]);
  const body = '<p style="font-size:13px">Hello Arietis team,</p><p style="font-size:13px">The escalations below are at or past our agreed turnaround (receipt confirmation within ' + s.receiptDays + " business days, resolution within " + s.resolutionDays + " business days). Please confirm receipt and provide a status update on each.</p>" +
    hTable(["Ticket", "Patient", "MRN", "Submitted", "Status", "Receipt", "Resolution"], rows) +
    '<p style="font-size:13px">Thank you,<br>' + escHtml(from || "MHS Billing") + "</p>";
  const cc = new Set<string>([s.billingInbox]);
  return { to: s.arietisEmail, cc: Array.from(cc).join(";"), subject: "MHS billing escalations — follow-up needed (" + list.length + ")", html: body, kind: "arietis" };
}

// ---------------------------------------------------------------------------
// Ops escalation review (shared with Arietis)
// ---------------------------------------------------------------------------
/** Open tickets in the ops escalation queue, follow-ups due first. */
function opsQueue(ts: Ticket[]): Ticket[] {
  return ts.filter((t) => t.opsEscalated && t.isOpen).sort((a, b) =>
    (a.followUpDue === null ? 1e9 : a.followUpDue) - (b.followUpDue === null ? 1e9 : b.followUpDue) ||
    a.topSev - b.topSev || (a.submitted || 0) - (b.submitted || 0));
}
/** Vendor-safe columns only: no internal work notes, review notes, owner emails or service-recovery flags. */
const ARIETIS_EXPORT_HEADERS: string[] = [
  "Ticket #", "Patient", "MRN", "Clinic", "Issue", "Amount", "Submitted", "Business days open", "Status",
  "Receipt SLA", "Resolution SLA", "Last Arietis reply", "Issue details", "Arietis commitment", "Follow-up due",
];
function arietisExportRows(ts: Ticket[]): string[][] {
  return ts.map((t) => [
    String(t.id), t.patient, t.mrn, t.clinic, t.taskTypes.join("; "), fmtMoney(t.amount), fmtDate(t.submitted), String(t.ageBD),
    t.status, slaText(t.receipt), slaText(t.resolution), fmtDateTime(t.lastReplyAt), t.notes, t.commitment, fmtDate(t.followUpDue),
  ]);
}
function buildArietisReview(ts: Ticket[], s: Settings, ctx: EvalCtx, from: string): Email {
  const rows = arietisExportRows(ts).map((r) => r.map((c) => escHtml(c)));
  const body = '<div style="font-family:Segoe UI,Arial,sans-serif;font-size:13px;color:' + EM.ink + '"><p>Hello Arietis team,</p>' +
    "<p>Below are the MHS patient billing escalations for review (" + ts.length + "). For each, please confirm the current status and the next step with a date.</p>" +
    hTable(ARIETIS_EXPORT_HEADERS, rows) +
    "<p>Turnaround agreement: receipt within " + s.receiptDays + " business days, resolution within " + s.resolutionDays + " business days of submission.</p>" +
    "<p>Thank you,<br>" + escHtml(from || "MHS Operations") + "</p></div>";
  return { to: s.arietisEmail, cc: s.billingInbox, subject: "MHS escalation review — " + ts.length + " tickets (" + fmtShort(ctx.now) + ")", html: body, kind: "arietis-review" };
}

// ---------------------------------------------------------------------------
// Dataverse mapping (Power Pages edition). One row per Master column.
// [Master column, Dataverse logical name, type, display name]
// types: int | text | memo | num | dt (date+time, user local) | date (date only) | bool
// ---------------------------------------------------------------------------
const DV_PREFIX = "mhs_";
const DV_TICKET_FIELDS: string[][] = [
  [COL.id, "mhs_formid", "int", "Ticket #"],
  [COL.start, "mhs_starttime", "dt", "Form start time"],
  [COL.submitted, "mhs_submittedon", "dt", "Submitted on"],
  [COL.email, "mhs_owneremail", "text", "Owner email"],
  [COL.name, "mhs_ownername", "text", "Owner name"],
  [COL.option, "mhs_requesttype", "text", "Request type"],
  [COL.dept, "mhs_department", "text", "Department"],
  [COL.clinic, "mhs_clinic", "text", "Clinic"],
  [COL.patient, "mhs_patient", "text", "Patient"],
  [COL.mrn, "mhs_mrn", "text", "MRN"],
  [COL.urgency, "mhs_urgency", "text", "Urgency"],
  [COL.source, "mhs_source", "text", "Source of inquiry"],
  [COL.taskType, "mhs_tasktype", "text", "Task type"],
  [COL.amount, "mhs_amount", "num", "Amount"],
  [COL.notes, "mhs_notes", "memo", "Notes"],
  [COL.attachments, "mhs_attachments", "memo", "Attachments"],
  [COL.status, "mhs_status", "text", "Status"],
  [COL.firstReplyDate, "mhs_firstreplydate", "date", "Date of first reply from Arietis"],
  [COL.outreachDate, "mhs_outreachdate", "date", "Date of patient outreach"],
  [COL.resolutionDate, "mhs_resolutiondate", "date", "Date of resolution"],
  [COL.falseVerif, "mhs_falseverification", "text", "False verification by Arietis"],
  [COL.serviceRecovery, "mhs_servicerecovery", "text", "Service recovery flag"],
  [COL.errorSource, "mhs_errorsource", "text", "Source of error"],
  [COL.outcome, "mhs_outcome", "text", "Outcome"],
  [COL.ehr, "mhs_ehrtask", "bool", "EHR task"],
  [COL.link, "mhs_formlink", "memo", "Form response link"],
  [COL.convId, "mhs_conversationid", "text", "Arietis conversation ID"],
  [COL.sentAt, "mhs_senttoarietisat", "dt", "Sent to Arietis at"],
  [COL.firstReplyAt, "mhs_firstreplyat", "dt", "Arietis first reply at"],
  [COL.lastReplyAt, "mhs_lastreplyat", "dt", "Arietis last reply at"],
  [COL.replyCount, "mhs_replycount", "int", "Arietis reply count"],
  [COL.cc, "mhs_cc", "text", "CC"],
  [COL.updBy, "mhs_lastupdatedby", "text", "Last updated by"],
  [COL.updAt, "mhs_lastupdatedat", "dt", "Last updated at"],
  [COL.rawSynced, "mhs_formsyncedat", "dt", "Form synced at"],
  [COL.workNotes, "mhs_worknotes", "memo", "Work notes"],
  [COL.opsFlag, "mhs_opsescalation", "text", "Ops escalation"],
  [COL.opsReason, "mhs_opsreason", "text", "Ops escalation reason"],
  [COL.opsBy, "mhs_opsescalatedby", "text", "Ops escalated by"],
  [COL.opsAt, "mhs_opsescalatedat", "dt", "Ops escalated at"],
  [COL.commitment, "mhs_arietiscommitment", "memo", "Arietis commitment"],
  [COL.followUpDue, "mhs_followupdue", "date", "Follow-up due"],
  [COL.reviewNotes, "mhs_reviewnotes", "memo", "Review notes"],
  [COL.reviewedAt, "mhs_lastreviewedat", "dt", "Last reviewed at"],
  [COL.ownerEmail, "mhs_assignedto", "text", "Assigned to (email)"],
  [COL.closedAt, "mhs_closedat", "dt", "Closed at"],
  [COL.followUpAt, "mhs_followupat", "date", "Last follow-up"],
];
const DV_ACTIVITY_FIELDS: string[][] = [
  ["at", "mhs_at", "dt", "When"], ["id", "mhs_ticketnumber", "int", "Ticket #"], ["field", "mhs_name", "text", "Field"],
  ["oldVal", "mhs_oldvalue", "memo", "Old value"], ["newVal", "mhs_newvalue", "memo", "New value"],
  ["by", "mhs_by", "text", "By"], ["source", "mhs_source", "text", "Source"],
];

/** US Pacific offset (minutes from UTC) for a UTC instant — used where no browser clock exists (Office Scripts). */
function pacificOffsetMin(utcMs: number): number {
  const y = new Date(utcMs).getUTCFullYear();
  const nthSunday = (month: number, n: number): number => {
    const first = new Date(Date.UTC(y, month, 1)).getUTCDay();
    return 1 + ((7 - first) % 7) + (n - 1) * 7;
  };
  const start = Date.UTC(y, 2, nthSunday(2, 2), 10); // 2nd Sun Mar, 2am PST = 10:00 UTC
  const end = Date.UTC(y, 10, nthSunday(10, 1), 9);   // 1st Sun Nov, 2am PDT = 09:00 UTC
  return utcMs >= start && utcMs < end ? -420 : -480;
}
/** Dataverse UTC ISO -> local wall-clock serial. offsetFn gives minutes east of UTC for that instant. */
function utcIsoToSerial(iso: string, offsetFn: (ms: number) => number): number | null {
  const ms = Date.parse(iso);
  if (isNaN(ms)) return null;
  return (ms + offsetFn(ms) * 60000) / 86400000 + 25569;
}
/** Local wall-clock serial -> UTC ISO for Dataverse. */
function serialToUtcIso(serial: number, offsetFn: (ms: number) => number): string {
  const wall = Math.round((serial - 25569) * 86400000);
  let ms = wall - offsetFn(wall) * 60000;
  ms = wall - offsetFn(ms) * 60000; // re-evaluate at the real instant (DST edges)
  return new Date(ms).toISOString().replace(/\.\d{3}Z$/, "Z");
}
function serialToIsoDate(serial: number): string {
  const d = serialToDate(Math.floor(serial));
  return d.getUTCFullYear() + "-" + pad2(d.getUTCMonth() + 1) + "-" + pad2(d.getUTCDate());
}
/** Converts one Dataverse value (as returned by the Web API) into the Master cell value. */
function dvToCell(type: string, v: string | number | boolean | null | undefined, offsetFn: (ms: number) => number): Cell {
  if (v === null || v === undefined || v === "") return type === "bool" ? false : "";
  if (type === "dt") { const s = utcIsoToSerial(String(v), offsetFn); return s === null ? "" : s; }
  if (type === "date") { const s = toSerial(String(v).slice(0, 10)); return s === null ? "" : s; }
  if (type === "bool") return v === true || v === "true";
  if (type === "int" || type === "num") return typeof v === "number" ? v : (toNum(String(v)) === null ? "" : (toNum(String(v)) as number));
  return String(v);
}
/** Converts a Master cell value into the Dataverse value to PATCH. Blank -> null. */
function cellToDv(type: string, v: Cell, offsetFn: (ms: number) => number): string | number | boolean | null {
  if (type === "bool") return toBool(v);
  if (isBlank(v)) return null;
  if (type === "dt") { const s = toSerial(v); return s === null ? null : serialToUtcIso(s, offsetFn); }
  if (type === "date") { const s = toSerial(v); return s === null ? null : serialToIsoDate(s); }
  if (type === "int") { const n = toNum(v); return n === null ? null : Math.round(n); }
  if (type === "num") return toNum(v);
  return norm(v);
}
interface DvRecord {
  [key: string]: string | number | boolean | null;
}
/** Dataverse ticket records -> { headers, rows } in Master shape (so every view/rule works unchanged). */
function dvTicketsToTable(recs: DvRecord[], offsetFn: (ms: number) => number): { headers: string[]; rows: Cell[][]; ids: string[] } {
  const headers = DV_TICKET_FIELDS.map((f) => f[0]);
  const rows: Cell[][] = [];
  const ids: string[] = [];
  for (const r of recs) {
    rows.push(DV_TICKET_FIELDS.map((f) => dvToCell(f[2], r[f[1]], offsetFn)));
    ids.push(String(r["mhs_ticketid"] || ""));
  }
  return { headers: headers, rows: rows, ids: ids };
}
function dvActivityRows(recs: DvRecord[], offsetFn: (ms: number) => number): ActivityRow[] {
  return recs.map((r) => ({
    at: (dvToCell("dt", r["mhs_at"], offsetFn) as number) || 0,
    id: toNum(dvToCell("int", r["mhs_ticketnumber"], offsetFn)) || 0,
    field: norm(r["mhs_name"] as string), oldVal: norm(r["mhs_oldvalue"] as string), newVal: norm(r["mhs_newvalue"] as string),
    by: norm(r["mhs_by"] as string), source: norm(r["mhs_source"] as string),
  }));
}
function dvField(col: string): string[] | null {
  for (const f of DV_TICKET_FIELDS) if (f[0] === col) return f;
  return null;
}

// ---------------------------------------------------------------------------
// Derived Master columns (written by the add-in sync and the Automation script)
// ---------------------------------------------------------------------------
function isoDay(serial: number | null): string {
  return serial === null ? "" : serialToIsoDate(serial);
}
/** Confirmation Due / Resolution Due / Escalation Flag / Reminder Flags for one evaluated ticket. */
function derivedColumns(t: Ticket): string[][] {
  const esc: string[] = [];
  const rem: string[] = [];
  if (t.isOpen) {
    for (const a of t.actions) (a.sev === 1 ? esc : rem).push(a.text);
    if (t.urgency === "Critical") esc.push("Critical priority");
    if (t.opsEscalated) esc.push("Ops escalation");
  }
  return [
    [COL.confirmDue, isoDay(t.receipt.due)],
    [COL.resolveDue, isoDay(t.resolution.due)],
    [COL.escFlag, esc.join("; ")],
    [COL.reminderFlags, rem.join("; ")],
  ];
}

// ---------------------------------------------------------------------------
// Clinic AM / PM recap (clinic help inboxes; includes patient code + MRN by design)
// AM: all open tickets + "Needs follow-up" (receipt past SLA) + "Resolution overdue" (resolution past SLA).
// PM: tickets resolved/closed on that local day.
// ---------------------------------------------------------------------------
interface ClinicRecap {
  clinic: string;
  subject: string;
  count: number;
  needsFollowUp: number;
  resolutionOverdue: number;
  html: string;
}
function recapTable(rows: Ticket[]): string {
  if (!rows.length) return "<p>No tickets in this section.</p>";
  const th = (h: string): string => '<th style="text-align:left;border:1px solid #c4d2ce;padding:8px">' + h + "</th>";
  const td = (v: string): string => '<td style="border:1px solid #c4d2ce;padding:8px">' + escHtml(v) + "</td>";
  return '<table style="border-collapse:collapse;width:100%"><thead><tr>' + ["Date Opened", "MRN", "Patient", "Ticket Number", "Reply Status", "Resolution Status"].map(th).join("") + "</tr></thead><tbody>" +
    rows.map((t) => "<tr>" + [
      fmtDate(t.submitted), t.mrn, t.patient, "#" + t.id,
      t.receipt.at !== null ? "Confirmed " + fmtDate(t.receipt.at) : t.receipt.state === "met-nodate" ? "Confirmed" : t.receipt.state === "breached" ? "Awaiting reply — " + t.receipt.bdOver + " BD past SLA" : "Awaiting reply",
      t.isOpen ? t.status + (t.resolution.state === "breached" ? " — " + t.resolution.bdOver + " BD past SLA" : "") : t.status,
    ].map(td).join("") + "</tr>").join("") + "</tbody></table>";
}
function buildClinicRecap(tickets: Ticket[], clinic: string, mode: string, ctx: EvalCtx, workbookUrl: string): ClinicRecap {
  const today = Math.floor(ctx.now);
  const mine = tickets.filter((t) => t.clinic === clinic);
  const closedDay = (t: Ticket): number | null => (t.closedAt !== null ? Math.floor(t.closedAt) : t.resolutionDate !== null ? Math.floor(t.resolutionDate) : null);
  const selected = mode === "pm" ? mine.filter((t) => !t.isOpen && closedDay(t) === today) : sortForAttention(mine.filter((t) => t.isOpen));
  const needs = selected.filter((t) => t.isOpen && t.receipt.state === "breached");
  const overdue = selected.filter((t) => t.isOpen && t.resolution.state === "breached");
  const safeUrl = /^https:\/\//i.test(workbookUrl || "") ? workbookUrl : "";
  const dayIso = serialToIsoDate(today);
  const intro = mode === "pm" ? "Please see the below billing tickets resolved or closed today." : "Please see the below open billing tickets";
  const html = '<div style="font:14px Arial;color:#203d42"><p>' + intro + "</p><h2>" + escHtml(clinic) + " · " + dayIso + "</h2>" + recapTable(selected) +
    (mode === "pm" ? "" : "<h3>Needs follow-up — receipt not confirmed within " + ctx.receiptDays + " business days</h3>" + recapTable(needs) +
      "<h3>Resolution Status Overdue — not resolved within " + ctx.resolutionDays + " business days</h3>" + recapTable(overdue)) +
    (safeUrl ? '<p><a href="' + escHtml(safeUrl) + '" style="background:#173f43;color:white;padding:10px 16px;text-decoration:none">Update tickets — open workbook and add-in</a></p>' : "") +
    "<p>SLA: " + ctx.receiptDays + " business days for receipt confirmation; " + ctx.resolutionDays + " business days for resolution, counted from submission. Weekends and listed holidays excluded.</p></div>";
  return { clinic: clinic, subject: clinic + " — " + (mode === "pm" ? "End-of-day billing recap" : "AM open billing tickets") + " — " + dayIso, count: selected.length, needsFollowUp: needs.length, resolutionOverdue: overdue.length, html: html };
}
```
