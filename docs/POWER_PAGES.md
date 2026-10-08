# MHS Billing Tickets — Power Pages edition (Dataverse)

The ticket desk runs as a web page at **testbilling.powerappsportals.com**, with no Excel involved. Tickets live in Dataverse. The Microsoft Form stays the intake, and a flow copies each submission into Dataverse.

```
Microsoft Form ──flow──► Dataverse: Billing Tickets ◄── Power Pages site (this app, signed-in staff)
billing@ inbox ──flow──►      ▲  ▲                      Ticket Activity (audit), Settings, Holidays, Reference Options
daily/weekly digest flow ─────┘  └── send-to-Arietis flow (Conversation ID)
```

---

## 1. Create the tables and migrate the 71 tickets (one time, about 10 minutes)

On a PC with Node 18 or newer, signed in as someone with **System Customizer** or **System Administrator** in the Power Pages environment:

```powershell
cd mhs-billing-excel-addin
npm install
node scripts/dataverse-setup.mjs --env https://<your-org>.crm.dynamics.com --data migration-data.json
```

- **Finding the environment URL:** Power Pages → ⋯ → **Admin center**, or Power Apps → ⚙ → **Session details** → *Instance url*.
- **Signing in:** you'll get a device code. Open the link, paste the code, and sign in.
- **What it creates:**
  - Publisher **mhs** and solution **Billing Escalations**.
  - The tables **Billing Tickets**, **Ticket Activity**, **Escalation Settings**, **Holidays** and **Reference Options**, with every column, plus a unique key on *Ticket #*.
  - Seed data: the default settings, the REF dropdown values, and all tickets from the Master tab.
  - `pages/config.js`.
- **Safe to re-run:** anything that already exists is skipped. Add `--dry-run` to preview without changing anything.
- **If sign-in is blocked:** your tenant may not allow Microsoft's public sign-in app. Ask IT for an app registration with *Dynamics CRM user_impersonation* and pass `--clientId <id>`.
- **Keep `migration-data.json` private.** It holds patient codes and MRNs. Delete it after the import, and never commit it.

Then run `npm run build`, which produces `dist-pages/`.

## 2. Configure the site (Power Pages design studio → **Set up**, plus the **Power Pages Management** app)

1. **Authentication:** keep Microsoft Entra ID sign-in and keep **Site visibility = Private**. Only signed-in MHS staff can reach it. Arietis never gets access; they see escalations only through the Ops review screen or the export.
2. **Web role:** use the built-in **Authenticated Users** role, or create a **Billing Staff** role and assign people to it.
3. **Table permissions** (Set up → Security → Table permissions), all with **Global access** and assigned to that role:

   | Table | Read | Write | Create |
   |---|:-:|:-:|:-:|
   | Billing Tickets | ✓ | ✓ | |
   | Ticket Activity | ✓ | | ✓ |
   | Escalation Settings | ✓ | | |
   | Holidays | ✓ | | |
   | Reference Options | ✓ | | |

4. **Site settings** for the Web API, created in the Power Pages Management app → Site Settings:

   | Name | Value |
   |---|---|
   | `Webapi/mhs_ticket/enabled` | `true` |
   | `Webapi/mhs_ticket/fields` | `mhs_ticketid,mhs_formid,mhs_starttime,mhs_submittedon,mhs_owneremail,mhs_ownername,mhs_requesttype,mhs_department,mhs_clinic,mhs_patient,mhs_mrn,mhs_urgency,mhs_source,mhs_tasktype,mhs_amount,mhs_notes,mhs_attachments,mhs_status,mhs_firstreplydate,mhs_outreachdate,mhs_resolutiondate,mhs_falseverification,mhs_servicerecovery,mhs_errorsource,mhs_outcome,mhs_ehrtask,mhs_formlink,mhs_conversationid,mhs_senttoarietisat,mhs_firstreplyat,mhs_lastreplyat,mhs_replycount,mhs_cc,mhs_lastupdatedby,mhs_lastupdatedat,mhs_formsyncedat,mhs_worknotes,mhs_opsescalation,mhs_opsreason,mhs_opsescalatedby,mhs_opsescalatedat,mhs_arietiscommitment,mhs_followupdue,mhs_reviewnotes,mhs_lastreviewedat` |
   | `Webapi/mhs_ticketactivity/enabled` | `true` |
   | `Webapi/mhs_ticketactivity/fields` | `mhs_at,mhs_ticketnumber,mhs_name,mhs_oldvalue,mhs_newvalue,mhs_by,mhs_source` |
   | `Webapi/mhs_setting/enabled` | `true` |
   | `Webapi/mhs_setting/fields` | `mhs_name,mhs_value` |
   | `Webapi/mhs_holiday/enabled` | `true` |
   | `Webapi/mhs_holiday/fields` | `mhs_name,mhs_date` |
   | `Webapi/mhs_refoption/enabled` | `true` |
   | `Webapi/mhs_refoption/fields` | `mhs_name,mhs_category,mhs_email,mhs_sort` |

5. If uploading fails on `.js` files, open Power Platform admin center → the environment → **Settings → Product → Privacy + Security → Blocked attachments** and remove `js`.

## 3. Upload the page

Install the CLI with `winget install Microsoft.PowerAppsCLI`, then run:

```powershell
pac auth create --environment https://<your-org>.crm.dynamics.com
pac pages upload-code-site --rootPath . --compiledPath ./dist-pages --siteName "Billing Escalations"
```

Next, go to Power Pages → **Inactive sites**, reactivate the site, and pick the `testbilling` address if it's free. Re-run the upload after any change.

**Check that it works:** sign in at the site. Your name should appear in the top-right chip; it comes from your Microsoft sign-in. **Mine** shows your tickets because your sign-in email matches the email the Form recorded. If your sign-in address differs from your Forms address, click the name chip to see which one is in use.

---

## 4. Power Automate flows (all standard connectors: Forms, Dataverse, Office 365 Outlook)

In every flow, turn on **⋯ → Settings → Concurrency control → Degree of parallelism = 1** on the trigger.

### Flow 1 — Intake: Form → Billing Tickets

Trigger: **When a new response is submitted** (Patient Billing Escalation Form). Then **Get response details**.

1. **Dataverse › List rows**
   - Table: *Billing Tickets*
   - Filter rows: `mhs_formid eq @{triggerOutputs()?['body/resourceData/responseId']}`
   - Row count: `1`
2. **Condition:** `length(outputs('List_rows')?['body/value'])` is equal to `0`.
   - **Yes → Dataverse › Add a new row** (*Billing Tickets*). Map each Form answer to the column in the table below. Also set:
     - *Ticket #* = Response Id
     - *Title* = `#@{responseId} @{Patient answer}`
     - *Owner email* = Responder's email
     - *Submitted on* = Submission time
     - *Status* = the Status answer, or `MHS - Submitted to Arietis` if blank
     - *Last updated by* = `form`
   - **No (an edit made through the Form's edit link) → Dataverse › Update a row**, with Row ID `first(outputs('List_rows')?['body/value'])?['mhs_ticketid']`. For each answer, use `if(empty(<answer>), null, <answer>)`. A null value leaves the column unchanged, so an edit can never blank out a field. Also add a **Ticket Activity** row with *Field* = `Form edit` and *By* = the responder's email.

| Form question | Dataverse column |
|---|---|
| Id | `mhs_formid` |
| Start time | `mhs_starttime` |
| Completion time | `mhs_submittedon` |
| Email | `mhs_owneremail` |
| Name | `mhs_ownername` |
| Choose an option: | `mhs_requesttype` |
| Department Submitting Ticket: | `mhs_department` |
| Patient Clinic: | `mhs_clinic` |
| Patient: | `mhs_patient` |
| MRN: | `mhs_mrn` |
| Urgency Level: | `mhs_urgency` |
| Source of Inquiry: | `mhs_source` |
| Task Type: | `mhs_tasktype` |
| Amount: | `mhs_amount` |
| Notes / Encounter# / Other relevant addition information: | `mhs_notes` |
| Please add any attachments here: | `mhs_attachments` |
| Status: | `mhs_status` |
| Date of first reply from Arietis: | `mhs_firstreplydate` |
| Date of Patient Outreach (if applicable): | `mhs_outreachdate` |
| Date of Resolution: | `mhs_resolutiondate` |
| Was false verification of balance/charge given by Arietis? | `mhs_falseverification` |
| Patient Feedback/Service Recovery Flag: | `mhs_servicerecovery` |
| Source of Error: | `mhs_errorsource` |
| Outcome: | `mhs_outcome` |

### Flow 2 — Save the Arietis Conversation ID (your existing send-to-Arietis flow)

After the step that gets the sent email's Conversation ID:
1. **List rows**: *Billing Tickets*, filter `mhs_formid eq <Response Id>`.
2. **Update a row** with Row ID `first(...)?['mhs_ticketid']`:
   - *Arietis conversation ID* = the Conversation ID
   - *Sent to Arietis at* = `utcNow()`

Recommended: start the email subject with `[BE-<Response Id>]`. If Arietis replies in a new thread, Flow 3 can still match it.

### Flow 3 — Arietis reply tracker

Trigger: **When a new email arrives in a shared mailbox (V2)**
- Mailbox: `billing@mindfulhealthsolutions.com`
- From: `patientbilling@arietishealth.com`
- Trigger condition: `@contains(toLower(triggerOutputs()?['body/from']),'patientbilling@arietishealth.com')`

1. **List rows** (*Billing Tickets*) with filter `mhs_conversationid eq '@{triggerOutputs()?['body/conversationId']}'`. If that returns nothing, run a second **List rows** with `mhs_formid eq @{if(contains(triggerOutputs()?['body/subject'],'[BE-'), int(first(split(last(split(triggerOutputs()?['body/subject'],'[BE-')),']'))), -1)}`.
2. Store the matching row in a **Compose** step named `Ticket`. **Condition:** the ticket exists **and** `not(or(startsWith(outputs('Ticket')?['mhs_status'],'Closed'), startsWith(outputs('Ticket')?['mhs_status'],'Resolved')))`.
3. **Update a row:**
   - *Arietis first reply at* = `if(empty(outputs('Ticket')?['mhs_firstreplyat']), triggerOutputs()?['body/receivedDateTime'], outputs('Ticket')?['mhs_firstreplyat'])`
   - *Date of first reply from Arietis* = `if(empty(outputs('Ticket')?['mhs_firstreplydate']), formatDateTime(convertFromUtc(triggerOutputs()?['body/receivedDateTime'],'Pacific Standard Time'),'yyyy-MM-dd'), outputs('Ticket')?['mhs_firstreplydate'])`
   - *Arietis last reply at* = `triggerOutputs()?['body/receivedDateTime']`
   - *Arietis reply count* = `add(coalesce(outputs('Ticket')?['mhs_replycount'],0),1)`
   - *Arietis conversation ID* = `coalesce(outputs('Ticket')?['mhs_conversationid'], triggerOutputs()?['body/conversationId'])`
4. **Add a new row** (*Ticket Activity*):
   - *Field* = `Arietis reply`
   - *Ticket #* = `outputs('Ticket')?['mhs_formid']`
   - *By* = sender
   - *Source* = `Inbox reply`
   - *When* = `utcNow()`

Once a ticket is Closed or Resolved, replies stop being logged, so *last reply* means the last reply before closing.

### Flows 4 & 5 — Daily (8:00 and 17:00, weekdays) and weekly (Fri 16:00) emails

The email logic is the same code the page uses (`src/rules.ts`), packaged as one Office Script. Office Scripts need a workbook to run in, so:
1. Create a blank workbook **Billing Escalations Engine.xlsx** in your OneDrive.
2. Open it, choose **Automate → New script**, paste in `office-scripts/DigestsDv.ts`, and save the script as `DigestsDv`.

Each flow works the same way. Only the trigger and the mode differ.

1. **Recurrence:**
   - Daily flow: weekdays, hours `8,17`, Pacific time.
   - Weekly flow: Friday, hour `16`.
2. **Dataverse › List rows** ×5:
   - *Billing Tickets* (all rows)
   - *Ticket Activity*, filtered `mhs_at ge @{addDays(utcNow(),-8)}`
   - *Escalation Settings*
   - *Holidays*
   - *Reference Options*
3. **Excel Online (Business) › Run script**: file *Billing Escalations Engine.xlsx*, script `DigestsDv`.
   - mode:
     - Daily flow: `if(less(int(formatDateTime(convertFromUtc(utcNow(),'Pacific Standard Time'),'HH')),12),'open','close')`
     - Weekly flow: `weekly`
   - nowLocal: `formatDateTime(convertFromUtc(utcNow(),'Pacific Standard Time'),'yyyy-MM-ddTHH:mm:ss')`
   - ticketsJson: `string(outputs('List_tickets')?['body/value'])`. Pass the other four lists the same way: activityJson, settingsJson, holidaysJson, refJson.
4. **Condition:** `outputs('Run_script')?['body/result/skip']` is false → **Apply to each** `body/result/emails` → **Send an email from a shared mailbox (V2)**, from billing@:
   - To = `item()?['to']`
   - CC = `item()?['cc']`
   - Subject = `item()?['subject']`
   - Body = `item()?['html']`

Who gets what: each owner, plus anyone CC'd on a ticket, gets a personalized start-of-day, end-of-day and weekly email. Leadership and clinic emails depend on the **Escalation Settings** rows (*LeadershipRecipients*, *ClinicDailyDigest*, *ArietisDailyFollowup*).

## Holidays
No holidays are seeded (the same rule as the Excel edition: holidays are never invented). Add your organization's dates as rows in **Holidays** (`mhs_name`, `mhs_date`). Every business-day calculation skips them.
