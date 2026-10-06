# MHS Billing Tickets — Excel add-in

Built for the supplied workbook: Master, REF, and Raw Data. No patient records are bundled or published. The original XLSX is unchanged to retain its Forms connection and metadata.

## Install
1. Keep the original Forms-linked workbook in its existing Microsoft 365 location. Uploading a downloaded copy does not automatically restore the original Forms live connection.
2. Download manifest.xml from https://brendannyates.github.io/mhs-billing-excel-addin/manifest.xml . The GitHub repository and Pages hosting are already configured.
3. Sideload the manifest using Excel's add-in upload option, or have Microsoft 365 IT deploy it centrally. Tenant policy can restrict sideloading.
4. Open the original linked workbook and launch MHS Billing Tickets.
5. Launch the add-in, choose **Initialize / refresh** once. It adds operational columns to Master, a hidden _SyncState ledger and Settings. Raw Data and OfficeForms.Table remain untouched. Then choose **Sync submissions**.
6. Configure name, email, role and assigned clinics in the add-in Settings. These are filters, not authentication or row-level security. Anyone with workbook access can read all its data.
7. Populate actual vendor-observed holidays in Settings column D. An empty list means weekdays only; holidays are not invented. SLA dates count from the day AFTER vendor submission, due by business close on the 2nd/5th working day. Enter Submitted to Vendor At as a Pacific YYYY-MM-DD date until outbound-message capture is connected.

## Clinic Dashboard and profile
The add-in opens to Clinic Dashboard. Configure name, email, role and, optionally, assigned clinics in Settings. The profile is saved in browser/device local storage; it contains no ticket or patient data. Role is descriptive; workbook permissions govern access. Dashboard defaults to assigned clinics, or to all clinics when none are picked. It can filter My Tickets (tickets you submitted, matched on the Forms Email even if Owner Email is reassigned), a specific clinic, exact MRN or all clinics. Status pills show Unresolved, Awaiting Reply, Past Due, Needs Follow-up, Resolution Overdue and All statuses. Update status opens that ticket's Master editor.

Auto-sync checks Raw Data every 30 seconds while the Excel task pane is active. It pauses for a busy operation, active input, open editor or Settings form. Closing/suspending the add-in stops the timer; unattended synchronization still requires the Microsoft 365 flow. Document vendor/staff follow-up in Follow Up At; the Past Due clock also recognizes first/last reply or patient outreach dates.

Formal SLA: 2 business days for receipt and 5 for resolution. Reminder thresholds are separate: receipt unconfirmed beyond 3 business days from ticket entry; no follow-up beyond 5 business days from latest documented follow-up (or entry); resolution incomplete beyond 6 business days from vendor submission (or entry if unrecorded). Weekends/Settings holidays do not advance these clocks.

Clinic Recaps generates copyable AM and end-of-day email templates. AM includes all open tickets for the selected clinic, required patient/MRN/date/status fields, Needs follow-up and Resolution Status Overdue sections. PM includes resolved/closed tickets from that Pacific calendar day. Set the cloud workbook URL in Settings for a general Update tickets link; email cannot reliably deep-link directly into a specific Excel add-in ticket. Recipients come from REF clinic emails. Templates are previews; no email is sent from the add-in.

All ticket edits write to Master. Forms writes to Raw Data through Microsoft's existing connection. The add-in does not scrape Forms or independently export responses. Microsoft 365 flows handle unattended ingestion and sending; see AUTOMATION.md.

## Merge policy
Join by Forms Id. Blank source values never erase Master. Changed nonblank source values update a field only when Master still matches its last observed source value. Initial populated disagreements and simultaneous edits preserve Master and appear as conflicts. Explicit add-in edits mark a field as Master-owned in _SyncState, including intentional clears. Staff can review an imported disagreement in Raw Data and choose the intended value in the Master editor. Direct sheet edits are detectable when their value differs from the baseline, but intentional blank clears should use the add-in.

No row is removed because a response disappeared. Raw Data response 31's missing clinic/MRN/notes remain preserved if already present on Master. Fields already lost from BOTH tabs cannot be reconstructed from blanks.

Owners close tickets in the add-in after entering resolution date and completing inbox reconciliation. Forms close-status changes are flagged rather than automatically closing a ticket. Reopening needs a reviewed workflow. First/last vendor reply times are read-only in the add-in.

## Operational limits
Repository: https://github.com/brendannyates/mhs-billing-excel-addin . The manifest uses the corresponding GitHub Pages host. Microsoft 365 installation and automation connections are separate setup steps.

The embedded Forms page can require sign-in or be blocked by browser/tenant iframe policy; the Open form link is provided. Office runtime, live Forms ingestion, tenant deployment, mailbox access and flow execution require Microsoft 365 validation. No live inbox scan or email sending is enabled by this package.

Excel provides no transactional row locks. Serialize automated writers in one worker flow and avoid staff saves during a sync write. This package detects a changed edited field before saving but cannot guarantee atomicity against simultaneous coauthoring; use a transactional backend if strict concurrent-write guarantees are required.

The app uses no analytics or external ticket-data APIs. Local storage holds only your manually entered profile. Runtime ticket data stays in memory and is read/written using Office.js; the static host serves code. M365 permissions and organization approval govern its use with this workbook.

## Verify
Run `npm test` for merge, blank-preservation, conflict, zero-value and SLA tests. Before rollout, use a tenant test workbook: new response, edited response with blanks, leading-zero MRN, staff edit followed by sync, duplicate Id refusal, receipt holiday boundary, two replies in reversed arrival order, unmatched conversation, closed-ticket reply cutoff, and email recipient verification.

Sources: Microsoft Excel add-ins overview https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-overview ; Forms/Power Automate https://learn.microsoft.com/en-us/power-automate/forms/overview .
