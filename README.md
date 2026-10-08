# MHS Billing Tickets

A ticket desk for the **Patient Billing Escalation** Microsoft Form. It is one codebase with two editions:

| Edition | Runs in | Tickets live in | Guide |
|---|---|---|---|
| **Excel add-in** (live) | The Forms-linked workbook: task pane, plus a full-screen window | Master tab (Raw Data is read-only) | this README + [AUTOMATION.md](AUTOMATION.md) |
| **Power Pages site** | testbilling.powerappsportals.com | Dataverse | [docs/POWER_PAGES.md](docs/POWER_PAGES.md) |

No patient records are bundled or published. The static host serves code only. Ticket data is read and written inside the user's own Excel or Power Pages session.

## What's in it (v2)
- **My Tickets:** tickets you submitted (matched on the Forms Email, even after reassignment) plus tickets you're CC'd on. Filters: Open, Needs action, Closed.
- **Action:** the needs-action queue (see *SLA rules*), with one-click Arietis follow-up and owner reminder drafts.
- **Ops:** the escalation queue. It holds Critical tickets, tickets with service recovery, and tickets an ops leader flags; *De-escalate* clears a ticket.
  - **Review with Arietis:** a share-safe, one-ticket-at-a-time mode for calls. It records Arietis's commitment, a follow-up date and review notes.
  - **Export for Arietis:** vendor-safe columns only.
- **Clinic:** pick a clinic (your assigned clinics are marked ★) and see its tickets by stage. **Email clinic** sends the clinic summary.
- **Stats:** a clinic × stage table, on-time % for receipt and resolution, average business days, $ open, and breakdowns.
- **MRN:** look up by MRN, patient code or ticket #.
- **Recaps:** the AM (open) and PM (closed today) clinic email template. Copy it, or draft it to the clinic inbox from REF.
- **Activity:** the Activity Log. Every import, form edit, conflict and add-in edit, with who, when, old value and new value.
- **+ New:** the embedded Form.
- **Ticket rows** are one line: severity dot, ticket #, patient code, MRN, issue, status. Every list has an **MRN / patient / #** filter.
- **Ticket detail** shows the SLA timeline, actions, the update form, the ops escalation card, any form/Master conflict, the Arietis thread and the history.
- **Profile** (name, email, role, assigned clinics) is stored on the device. Assigned clinics are the default scope for Action, Ops and Stats.
- **Full screen** (⛶): opens the same app in a large window. The task pane stays open behind it and does all workbook reads and writes.

## Install (Excel add-in)
1. Keep the original Forms-linked workbook in its Microsoft 365 location. A downloaded copy loses the live Forms connection.
2. Manifest: https://brendannyates.github.io/mhs-billing-excel-addin/manifest.xml . It uses the same add-in ID as v1, so existing installs update in place.
3. Sideload it (Excel → Add-ins → Upload My Add-in), or have IT deploy it under **Microsoft 365 admin center → Integrated apps**. Tenant policy can restrict sideloading.
4. Open the workbook and click **Billing Tickets** on the Home tab. On first open it runs **Initialize / refresh** automatically. That step:
   - adds the operational columns to Master
   - adds a hidden `_SyncState` ledger
   - adds the Settings rows and the **Activity Log** sheet

   Raw Data and `OfficeForms.Table` are never touched.
5. **Who you are:** once single sign-on is set up ([docs/SSO_SETUP.md](docs/SSO_SETUP.md)), your name and work email come from your Microsoft 365 sign-in automatically. Until then the add-in asks once. Role and assigned clinics are optional picks. Identity filters and attributes edits; it isn't access control, because the workbook's SharePoint permissions decide who can see data.
6. Put your actual holidays in **Settings column D** (YYYY-MM-DD). An empty list means weekdays only; holidays are never invented. The **Workbook URL** (used by email links) is filled in automatically with the MHS-only workbook link.

## SLA rules
- **Clocks start at form submission** (`Completion time`).
  - **Receipt** is due 2 business days later.
  - **Resolution** is due 5 business days later.
  - Weekends and Settings holidays are skipped, and the due day itself is not overdue.
  - The day counts are set on the Settings sheet.
- **Receipt stops** at the first Arietis reply recorded by the inbox flow, or when status moves past *MHS - Submitted to Arietis*.
- **Resolution stops** at the date of resolution. Closing requires a resolution date; the add-in fills in today's date when you pick a closed status. Reopening is blocked.
- **Needs action:**
  - past SLA or due today (vendor)
  - Arietis replied but status not updated
  - clinic or patient blank
  - pending patient call with no outreach date
  - form/Master conflict to review
  - no activity for 3 business days
  - closed without resolution date, error source, outcome or false-verification
  - ops follow-up due or overdue
- **Written to Master by sync:** Confirmation Due, Resolution Due, Escalation Flag and Reminder Flags.

## Sync and merge policy
- **Timing:** sync runs every 30 seconds while the task pane is open. It pauses while you're editing, and there's also a Sync button. Unattended sync needs the Power Automate worker in AUTOMATION.md.
- **Matching:** rows join on the Forms `Id`. A new Id is copied to Master and logged as **Imported**.
- **Blanks:** a blank form value never erases Master. That's why response #31's clinic, MRN and notes survived its edit-link overwrite.
- **Changed form values:** a changed non-blank value updates Master only if Master still equals the last value seen from the form. Otherwise Master is kept and a **conflict** is flagged on the ticket and in the Activity Log.
- **Staff-owned fields:** status, the reply, outreach and resolution dates, false verification, service recovery, source of error, outcome and EHR task are filled from the form only while blank. After that they're never overwritten and never flagged.
- **Add-in edits** mark a field Master-owned in `_SyncState`. A save is refused if someone else changed that field since you opened the ticket.
- **First and last Arietis reply times** are set by the inbox flow and are read-only in the add-in.

## Operational limits
- **Concurrency:** Excel has no transactional row locks. Serialize automated writers through one worker flow, and avoid staff saves during a worker write. The add-in detects a changed field before saving, but can't guarantee atomicity against simultaneous co-authoring.
- **Embedded form:** it can require sign-in or be blocked by iframe policy. Use the **Open in browser** button instead.
- **Full screen:** needs Office Dialog API 1.2. If your Excel doesn't support it, the button is hidden.
- **Emails from the add-in** open as Outlook drafts; nothing is sent automatically from the add-in. Scheduled emails come from the flows.

## Repository layout
| Path | What |
|---|---|
| `index.html`, `app.js`, `style.css` | The app UI (shared by both editions) |
| `excel.js` | Workbook reads and writes: Initialize, sync (merge + `_SyncState`), save, Activity Log |
| `core.js` | Sync merge policy (repo v1, unchanged behavior) |
| `src/rules.ts` → `rules.js` | **Rules engine:** SLA clocks, actions, ops escalation, stats, recaps, email builders, Dataverse mapping |
| `src/office/*.ts` → `office-scripts/*.ts` | Power Automate scripts (generated; paste-ready) |
| `pages/`, `scripts/dataverse-setup.mjs` | Power Pages edition: Dataverse data source and the one-time table setup and migration |
| `demo.js` | Synthetic data when the page is opened outside Excel |
| `docs/` | Power Pages guide and the build spec |

## Develop and verify
```bash
npm run build   # rules.js, office-scripts/, dist-pages/ from src/ (no dependencies; Node 22.6+)
npm test        # merge policy, Automation and ClinicRecap scripts, rules engine
npm run serve   # http://localhost:8080 → demo mode
```
CI refuses to publish if the generated files are stale or any test fails.

**Before rollout, test in a tenant test workbook:**
- a new response, and an edited response with blanks
- a leading-zero MRN
- a staff edit followed by a sync
- duplicate Id refusal
- the receipt holiday boundary
- two replies arriving in reverse order, and an unmatched conversation
- the reply cutoff after a ticket closes
- email recipients

Sources: [Excel add-ins overview](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-overview) · [Office dialog API](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/dialog-api-in-office-add-ins) · [Forms + Power Automate](https://learn.microsoft.com/en-us/power-automate/forms/overview)
