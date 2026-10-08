---
name: billing-addin-change
description: Make any change to the MHS Billing Tickets Excel add-in (views, SLA rules, escalation logic, emails, Office Scripts) end to end, from edit to test to push, and tell Brendan what to re-paste or re-upload.
---

# Change the MHS Billing Tickets add-in

Use this skill for any feature, fix or rule change in `mhs-billing-excel-addin`. Read `CLAUDE.md` first; its hard rules always apply.

## 1. Pin down the request
- Restate the change in one or two sentences. Ask if anything is ambiguous, and don't guess:
  - which views it affects
  - whether it should apply to all clinics or only assigned ones
  - whether it changes an email
- A change to the SLA, escalation, merge policy or Raw Data handling needs Brendan's explicit OK on the exact rule before you code it.

## 2. Find where it lives
| Kind of change | File(s) |
|---|---|
| Screen, tab, filter, ticket row/detail, profile | `app.js`, `index.html`, `style.css` |
| Initialize, sync/merge, save, Activity Log, Settings sheet | `excel.js` (merge helpers in `core.js`) |
| SLA clocks, needs-action list, severities, ops escalation, stats, derived Master columns, recap/digest/Arietis email builders | `src/rules.ts` |
| Power Automate script entry points (sync, reply capture, digests, clinic recap) | `src/office/*.ts` |
| Ribbon button, icons, permissions, host URL | `manifest.xml` |

- **Never hand-edit the generated files** `rules.js` and `office-scripts/*.ts`; they're rebuilt from `src/`.
- **New Master columns:**
  - Add them to the column map in `src/rules.ts` (`EXTRA_COLS` for columns the add-in creates).
  - Initialize must stay idempotent.
  - Raw Data and `OfficeForms.Table` are never written.
- **New Settings keys:** add a row to `SETTING_DEFAULTS` in `excel.js` so Initialize creates them. Readers normalize labels, so `Receipt SLA business days` is matched ignoring spaces and punctuation.

## 3. Test
1. Add or update a test in `tests/`. Use synthetic data only, with no real MRNs or patient codes.
   - Rule changes go in `tests/rules.test.mjs`.
   - Script changes go in `tests/automation.test.mjs` or `clinic-recap.test.mjs`.
   - Merge changes go in `tests/core.test.mjs`.
2. Run `npm run build && npm test` and make sure every test passes.
3. For UI changes, run `npm run serve` and check http://localhost:8080 (demo mode, synthetic data) in a browser at both task-pane width (~350px) and full-screen width.

## 4. Ship
1. Commit with a plain-English message, then `git push origin main`.
2. Check that the **pages** GitHub Action succeeds (`gh run list --limit 1`). If it fails, fix it before reporting.

## 5. Tell Brendan what to do
End with a short note that covers only what applies:
- **What changed:** one or two lines.
- **Live:** "Live now. Close and reopen the Billing Tickets pane to see it."
- **Scripts to re-paste:** only if anything under `office-scripts/` changed. List each one, for example: "Re-paste `office-scripts/Digests.ts` into **Billing Digests** (Automate → open it → select all → paste → save)."
  - Script names: `Automation.ts` → Billing Automation, `ClinicRecap.ts` → Billing Clinic Recap, `Digests.ts` → Billing Digests.
  - Attach or link the exact file.
- **Manifest:** only if `manifest.xml` changed. "Everyone re-uploads manifest.xml (or IT updates it in Integrated apps)."
- **Flows:** only if a script's inputs or outputs changed. Say which flow step in `AUTOMATION.md` needs updating, and how.
- **Settings:** if a new Settings row was added, give its default and what to set it to.
