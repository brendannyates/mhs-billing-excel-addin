# Power Automate configuration

This is a runnable Office Script plus flow build instructions, not an exported/importable Power Platform solution. Connections, mailbox address, SharePoint/OneDrive workbook identifier, business hours and actual holiday calendar must be supplied in your tenant. No flows are activated here.

## Shared workbook setup
Initialize using the Excel add-in. Save `office-scripts/Automation.ts` as **Billing Automation** in Excel > Automate > New Script. Use Excel Online (Business) **Run script** against the original linked workbook, not a downloaded duplicate.

All mutation jobs must go through ONE serialized worker. Setting concurrency=1 on separate independent flows does not serialize those flows against one another. Use an M365 queue (e.g. a restricted SharePoint list) with Operation, Payload, LocalToday, Status, Error; an item-created worker flow with trigger concurrency=1 runs the script and marks completion. Restrict access to the queue and workbook. Prevent direct parallel Run script calls. The app writes must also be scheduled outside worker writes or operationally coordinated; Excel has no transaction lock shared with Office.js.

Pacific date expression for script `localToday`:
```
formatDateTime(convertTimeZone(utcNow(),'UTC','Pacific Standard Time'),'yyyy-MM-dd')
```

The script returns JSON as the Run script `result` string. Parse that string into an object. Shape:
```
{ "processed": 71, "unmatched": [], "digests": [
  { "email": "owner@example.com", "tickets": [
    { "Id": 31, "Patient Clinic:": "San Rafael - CA", "Status:": "MHS - Submitted to Arietis",
      "Confirmation Due": "2026-10-06", "Resolution Due": "2026-10-09",
      "Escalation Flag": "", "Sync Conflicts": "" }
  ] }
] }
```
The example above is illustrative, not a workbook extract.

## 1. Raw Data ingestion and Master reconciliation
Keep the existing native Forms-linked table `OfficeForms.Table` on Raw Data. Do not rename, replace, resize, clear, add calculated columns to, or repoint that Forms table. The add-in never writes to it.

Use Microsoft Forms **When a new response is submitted** > **Get response details** to create a sync queue job after the linked Raw Data row is visible. Add retry/backoff for sync latency. Also schedule reconciliation (e.g. every 15 minutes during business hours) for response edits; the new-response trigger is not an edit-response trigger.

If native Forms sync is not updating this workbook unattended, enable the Microsoft's supported Forms-to-Excel synchronization for its original cloud workbook. Verify with the workbook closed. Alternatively use a dedicated ingestion flow that maps response details into the original Raw Data table by Id; choose ONE writer and validate this with your tenant before changing native sync. Do not run a second append flow beside native Forms sync: it creates duplicates. Do not assume scheduled Master reconciliation can retrieve response changes absent from Raw Data.

Worker parameters: `operation=sync`, `localToday=<Pacific date>`, `repliesJson=[]`, `mode=open`. Repeat runs are idempotent on unchanged input. Merge preserves filled Master values against blanks. Review populated disagreements under Needs action. Never infer a ticket relationship from MRN alone.

## 2. Register outbound vendor conversation
Each ticket must have ONE dedicated vendor conversation before automatic reply matching works. Map outbound billing mail to Ticket Id when sent; capture Outlook/Graph conversationId back to **Vendor Conversation ID** on Master. Capture actual submission date into **Submitted to Vendor At** as a Pacific YYYY-MM-DD value; SLAs start on that date, NOT automatically on form completion. Keep original submission date for reopened/continued conversations unless policy explicitly establishes a new SLA.

If the sender action does not return conversationId, retrieve the sent message using its message identifier or a unique ticket token in its subject and bounded sent-time search, then store its conversationId after verifying exactly one match. This initialization may use the ticket token; subsequent Arietis matching uses conversationId. Do not assign one conversation to multiple tickets. The add-in lets a user enter an already verified conversationId.

## 3. Inbox reply capture
Trigger **Office 365 Outlook — When a new email arrives in a shared mailbox (V2)** for the actual MHS billing mailbox (address not supplied). Filter exact normalized From address `patientbilling@arietishealth.com`. Take the actual sender address property, not display name. Include all monitored mail folders or avoid inbox rules that bypass the monitored folder; add scheduled reconciliation for moved messages.

Queue `operation=reply`, and `repliesJson` with this shape:
```
[{"conversationId":"<Outlook conversationId>","sender":"patientbilling@arietishealth.com",
  "receivedAt":"2026-10-02T21:30:00Z","messageId":"<Outlook message id>"}]
```
The script uses the message received timestamp, not flow execution time. First reply=min(receivedAt), last reply=max(receivedAt), tolerating retries and out-of-order events. Exactly one ticket must match; zero/ambiguous matches return message ids in `unmatched` for a restricted review queue. No automatic MRN or subject fallback.

An Arietis message counts as receipt confirmation by default. If automated acknowledgments or unrelated replies should not count, add an explicit policy filter before queueing the message. Receipt does not imply resolution; only staff closes the ticket after review. Replies received after Closed At do not change the frozen history; a late-processed message received before closure can correct the history.

Before an owner closes, retrieve all messages in the mapped conversation from the MHS billing inbox, filter Arietis sender, enqueue reconciliation and wait for worker success. Then refresh the add-in and close the ticket. Until an authenticated close-request orchestration is configured, this is a required manual operational checkpoint: the add-in cannot itself scan the mailbox or guarantee a catch-up completed. A future closing flow should serialize reconciliation and closure in the same worker.

## 4. Personalized daily emails
Create recurrence flows for 09:00 and 17:00 **Pacific Standard Time**, Monday–Friday (or your actual operating hours). Skip dates found in Settings holiday column D. After a successful sync/reply catch-up, run `operation=digest`, `mode=open` or `close`, `localToday=<Pacific date>`.

For each item in parsed `digests`, **Create HTML table** from its tickets, then **Send an email (V2)** To=`email`. Use subject `Billing tickets — opening review — <date>` or `Billing tickets — closing review — <date>`. Include the restricted cloud workbook URL, overdue flags, confirmation/resolution deadlines, open tickets and tickets closed that local day. The script's digest output deliberately excludes patient names, MRNs, notes, amounts and attachment URLs. Verify addresses against your approved staff directory before sending. Owner Email is the recipient; it defaults to the submitter Email and can be reassigned in Master.

Use a delivery log keyed by `local-date + opening/closing + recipient` to avoid duplicate email on retries. Mark sent after success; uncertain send outcomes require review, not blind retransmission. Do not send empty digests. Flow failures should go to the designated operations owner, who is not specified yet.

## 5. Weekly recap and escalation
Friday at business close, after reconciliation: `operation=digest`, `mode=weekly`. Include all active tickets and closures in the previous seven calendar days. Use subject `Billing weekly recap — <date>`. The per-ticket flags identify overdue receipt, overdue resolution, critical priority, missing submission date and sync review.

Send owner recaps as above. For clinic escalation, group flagged tickets by Patient Clinic and look up ClinicName/ClinicEmail in REF (Table2). Validate recipients, then send the approved clinic escalation digest. Define any additional escalation manager separately; no recipient is guessed. Route unresolved conversation matches to operations rather than treating them as receipt confirmations.

## Acceptance before enabling
- Confirm Forms writes to Raw Data while workbook is closed and changed-response reconciliation sees updates.
- Verify response #31 retains Master clinic/MRN/notes when source blanks them.
- Verify zero-valued amounts import and leading-zero MRNs remain text.
- Verify populated staff and source disagreements flag rather than overwrite.
- Verify vendor submission on Friday, a configured Monday holiday, receipt deadline Wednesday and resolution deadline the following Monday.
- Verify conversation mapping, two reversed-order replies, retries, unmatched/ambiguous messages and cutoff at closure.
- Verify organization holidays, Pacific daylight-saving scheduling, approved owner recipients, digest contents and no duplicate sends.

Microsoft references:
https://learn.microsoft.com/en-us/power-automate/forms/overview
https://learn.microsoft.com/en-us/office/dev/scripts/develop/power-automate-integration
https://learn.microsoft.com/en-us/connectors/office365/
https://learn.microsoft.com/en-us/connectors/excelonlinebusiness/


## Clinic AM / PM recap flow (updated dashboard)
Save office-scripts/ClinicRecap.ts as **Billing Clinic Recap**. After successful serialized reconciliation, run this read-only script at AM/PM business times with localToday as Pacific YYYY-MM-DD, mode=am or pm, clinic blank for all REF clinics (or one exact ClinicName), and workbookUrl set to the restricted workbook URL (or Settings B6).

Parse result and loop over clinicDigests. Use email (REF ClinicEmail), subject, and html directly in Send an email (V2), with HTML enabled. The script skips weekends and configured holidays and excludes empty recaps. AM starts “Please see the below open billing tickets” and lists Date Opened, MRN, Patient, Ticket Number, Reply Status, Resolution Status. It excludes resolved/closed tickets and adds receipt follow-up beyond 3 working days and resolution-overdue beyond 6 working days. PM lists only tickets resolved/closed on that Pacific local day. These clinic templates contain patient/MRN data; send only through your approved Microsoft 365 workflow to verified REF clinic recipients. They are generated in the workbook context and are never stored on GitHub.

Use a distinct delivery-log key local-date + clinic + am/pm for retry deduplication. Keep the existing personalized daily/weekly owner flow separate; its compact digests now include Reminder Flags as well as formal SLA flags. Formal deadlines remain 2/5 business days and do not shift to the 3/6 reminder thresholds.

The Update tickets email button is a general workbook link. After opening the workbook, launch MHS Billing Tickets and use Update status on the required ticket. No per-ticket email deep link is claimed. Ticket entry date means Completion time, falling back to Start time. Past Due means more than 5 business days since the most recent Follow Up At, patient outreach, first/last vendor reply, or entry if none exists. Record staff follow-ups on Master via the add-in.
