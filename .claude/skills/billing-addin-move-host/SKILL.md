---
name: billing-addin-move-host
description: Move the MHS Billing Tickets repo and its GitHub Pages hosting to a different GitHub owner or org (for example the MHS org) and re-point the add-in manifest without breaking users.
---

# Move the add-in to a new GitHub owner

Use this skill when the repo is being, or has been, transferred, for example from brendannyates to the MHS GitHub org. The GitHub Pages address **does not redirect**, so the manifest must change and every user must re-upload it.

## Before the transfer (ask, don't assume)
1. Get the exact new owner or org name.
2. Confirm with Brendan or IT:
   - **GitHub Pages and Actions** are allowed in that org.
   - If the repo will be **private**, the org has a paid plan (Team or Enterprise), because Pages on private repos requires one.
   - The Claude GitHub app is approved for that org, if Claude sessions will push there.
3. Pick a cut-over time. Old installs stop loading once the old Pages site is gone.

## Transfer (Brendan or an org owner does this in the browser)
1. Open the repo and go to **Settings → Danger Zone → Transfer ownership**.
2. Pick the org and confirm by typing the repo name.

## Re-point (Claude does this)
1. Run `node scripts/set-host.mjs <new-owner> --dry-run` and review the files it lists.
2. Run `node scripts/set-host.mjs <new-owner>`. It rewrites every URL and bumps the manifest `<Version>`.
3. Run `git remote set-url origin https://github.com/<new-owner>/mhs-billing-excel-addin.git`.
4. Run `npm run build && npm test`, commit, and push to `main`.
5. In the new repo, set **Settings → Pages → Source: GitHub Actions**. Make sure the **pages** workflow succeeds, and that `https://<new-owner>.github.io/mhs-billing-excel-addin/manifest.xml` loads.

## Tell Brendan
- The new manifest URL, with the file attached.
- Everyone must:
  1. Remove the old add-in. Use the billing-addin-rollout steps; on desktop that means clearing the Wef cache.
  2. Upload the new manifest. Or IT replaces it once in Integrated apps.
- The Office Scripts, flows and workbook are **unaffected**; they live in the workbook.
- Update any bookmarks or docs that link to the old repo.
