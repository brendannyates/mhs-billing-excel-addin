# Working on this repo (instructions for Claude)

MHS Billing Tickets: an Excel add-in that serves as a ticket desk over the Forms-linked **Patient Billing Escalation Forms** workbook.

- Tickets are escalated to Arietis (patientbilling@arietishealth.com).
- The live edition is the **Excel add-in**, hosted on GitHub Pages.
- The Power Pages edition (`pages/`, `docs/POWER_PAGES.md`, `scripts/dataverse-setup.mjs`, `office-scripts/DigestsDv.ts`) is parked. Don't extend it unless asked.

## Hard rules
- **Never write to the `Raw Data` sheet or `OfficeForms.Table`.** Forms owns them.
  - All submissions are copied to the **Master** sheet.
  - All edits happen on Master.
- **No patient data in the repo.** Never commit:
  - workbook exports
  - `migration-data.json`
  - screenshots or fixtures with real MRNs or patient codes

  Tests use synthetic data only.
- **Don't change the merge policy without asking.** It lives in `core.js` and `excel.js` sync:
  - Blanks never erase Master.
  - A conflict is flagged rather than overwritten.
  - Staff-owned fields are fill-only-while-blank (`MASTER_OWNED_COLS`).
  - `NEVER_MERGE_COLS` are skipped.
- **SLA rules:**
  - Clocks start at `Completion time`.
  - Receipt is due +2 business days and resolution +5.
  - Each clock is due through the end of the due day.
  - Weekends are skipped, and so are holidays from Settings column D. Never invent holidays.
- **Don't change the manifest `<Id>`.** Changing it breaks existing installs.

## Where things live
| Change | Edit | Then |
|---|---|---|
| UI, views | `app.js`, `index.html`, `style.css` | none |
| Workbook reads, writes, sync | `excel.js`, `core.js` | none |
| SLA, actions, ops escalation, stats, emails | **`src/rules.ts`** (never hand-edit `rules.js`) | `npm run build` |
| Power Automate scripts | **`src/office/*.ts`** (never hand-edit `office-scripts/`) | `npm run build` |
| Ribbon, icons, permissions, host URL | `manifest.xml` | bump `<Version>` |

## Before every commit
```bash
npm run build && npm test
```
- CI (`.github/workflows/pages.yml`) refuses to publish if the generated files are stale or any test fails.
- Add or adjust tests in `tests/` for any rule change.
- `npm run serve` gives a demo mode that runs on synthetic data.

## Deploying
- **Push to `main`** and GitHub Pages publishes within about 1 minute. Users get the change when they reopen the task pane.
- **If `office-scripts/*.ts` changed,** tell the user exactly which scripts to re-paste in the workbook (Automate → open the script with the same name → replace all → save):
  - `Automation.ts` → **Billing Automation**
  - `ClinicRecap.ts` → **Billing Clinic Recap**
  - `Digests.ts` → **Billing Digests**

  Rule changes usually touch both the add-in and the scripts. Say so.
- **If `manifest.xml` changed,** tell the user that everyone must re-upload it, or IT must update it in Integrated apps.
- **If the GitHub owner or org changes,** update every URL in `manifest.xml`, `README.md` and the docs. The Pages address does not redirect.

## Working with Brendan
- Ask rather than assume when a request is ambiguous, and state any assumptions you make.
- Don't assume anyone's gender from a name.
- Reply in plain operational language: what changed, and what he needs to do (if anything).

## Skills in this repo
`.claude/skills/` holds three skills:
- `billing-addin-change`: make any change end to end.
- `billing-addin-rollout`: install, update, remove, troubleshoot and onboard.
- `billing-addin-move-host`: move to a new GitHub owner. It uses `scripts/set-host.mjs`.
