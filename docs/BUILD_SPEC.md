# BUILD SPEC — MHS Billing Tickets on Power Pages

> **For Claude:** this file is the brief; read it alongside the repo (§12 lives there). Build the app it describes as a **Power Pages code site** (a React + Vite + TypeScript single-page app) in VS Code, backed by Dataverse. Work through §10 *Build order* and check off §11 *Acceptance tests*. **§12 points to the tested reference implementation** (`src/rules.ts`) of every business rule. Port it as a module and don't re-derive the rules. Ask the owner before changing anything marked **LOCKED**.
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
The tested rules engine lives in the repo at **`src/rules.ts`** (covered by `tests/rules.test.mjs`). Read it from there; this spec deliberately no longer keeps a copy, because a pasted copy drifts out of date. It's TypeScript compatible with Office Scripts, so it has no imports, no `any` and no DOM. In the React app, export the functions from it. The `COL` names are the Excel Master headers; keep them, because the Dataverse mapping (`DV_TICKET_FIELDS`) is defined in terms of them.
