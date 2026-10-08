---
name: billing-addin-rollout
description: Walk Brendan or MHS staff through installing, updating, removing or troubleshooting the MHS Billing Tickets add-in, its Office Scripts and Power Automate flows, and onboarding a new user.
---

# Roll out and support MHS Billing Tickets

Use this skill for "how do I install / update / remove / fix" questions and for onboarding staff. Answer with short, numbered steps that include the exact menu names. Ask which situation applies (Excel on the web or desktop, who installed it) rather than assuming.

**The manifest URL** is the GitHub Pages URL in `manifest.xml`: `<SourceLocation>` minus `index.html`, plus `manifest.xml`. Read it from the file each time; don't rely on memory.

## Install (per person)
1. Open the **original Forms-linked workbook** from SharePoint or OneDrive. A downloaded copy loses the Forms link.
2. Go to **Home → Add-ins → My Add-ins → Upload My Add-in** and choose `manifest.xml`.
   - If that option is missing, the tenant blocks uploads. IT deploys it under **M365 admin center → Integrated apps** instead.
3. Click **Home → Billing Tickets**. The first open runs Initialize, which:
   - adds the Master columns, the Settings rows and the Activity Log sheet
   - never touches Raw Data
4. Fill in the profile: name, the email used in Forms, role, assigned clinics. The profile is a filter, not security.

## One-time workbook setup (owner)
On the Settings sheet, fill in:
- Workbook URL
- Holidays in column D (YYYY-MM-DD; leave blank for weekdays only)
- Leadership recipients
- Clinic daily digest and Arietis daily followup (Yes/No)
- Digest include patient IDs (default No)
- Receipt and resolution SLA business days (2 and 5)

## Office Scripts (owner)
1. In the workbook, go to **Automate → New Script**, replace all of the starter code with the file's contents, and rename it exactly:
   - `office-scripts/Automation.ts` → **Billing Automation**
   - `office-scripts/ClinicRecap.ts` → **Billing Clinic Recap**
   - `office-scripts/Digests.ts` → **Billing Digests**
2. To update a script, open the existing one, select all, paste, and save.
3. Replace an old Billing Automation **before** any scheduled flow runs; mismatched versions fight each other.

## Flows
`AUTOMATION.md` is the source of truth for triggers, inputs and the acceptance checklist. The flows are:
- Forms → sync
- send-to-Arietis capture
- inbox replies from patientbilling@arietishealth.com
- digests: open (weekdays 9:00), close (weekdays 17:00), weekly (Friday close)
- clinic recaps (AM and PM)

Recommend running the README "test in a tenant test workbook" list before turning the flows on for everyone.

## Updates
- **App changes** go live automatically after a push to `main`. Users close and reopen the pane, or Excel if it's cached.
- **Script changes:** re-paste only the changed scripts.
- **Manifest changes:** everyone re-uploads, or IT updates it once.

## Remove old add-ins
- **Excel on the web:** Home → Add-ins → More Add-ins → My Add-ins → hover → **…** → Remove.
- **Windows desktop (uploaded add-ins):**
  1. Close Excel.
  2. Open `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` and delete its contents. This clears the add-in cache, not data.
  3. Reopen Excel and re-upload the current manifest.
- **Shared-folder catalog:** File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs → remove the path.
- **IT-deployed:** IT removes it in Integrated apps.
- The old "Billing Escalations" and "Billing Portal" add-ins should be removed. A v1 "MHS Billing Tickets" install updates in place because it has the same ID.

## Troubleshooting
| Symptom | Check |
|---|---|
| Pane is blank or shows the old version | Close Excel fully. On desktop, clear the Wef cache. Check that the latest **pages** Action succeeded. |
| "Not in the Forms workbook" / no Master | Wrong file, or a downloaded copy. Open the SharePoint original. |
| Conflicts flagged on a ticket | Expected when the form and Master disagree. Open the ticket and pick the value; the Activity Log shows both. |
| Save refused | Someone else changed that field since the ticket was opened. Reopen and retry. |
| Emails missing or wrong people | Settings recipients and the Yes/No toggles. Check the flow run history. Check that the script was re-pasted after the last change. |
| SLA dates look off | Settings SLA days and the holidays in column D. Clocks start at Completion time. |
| Full-screen button missing | That Excel build lacks Dialog API 1.2. Use the pane. |
| Embedded form won't load | Use **Open in browser**. Iframe or sign-in policy blocks it. |

**Privacy:** never ask for or accept workbook exports or screenshots with MRNs or patient names in this chat unless MHS has approved this Claude account for PHI. Ask for a ticket # instead.
