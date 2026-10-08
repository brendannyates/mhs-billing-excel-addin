// Re-point the add-in at a new GitHub owner (user or org) after a repo transfer.
// Usage: node scripts/set-host.mjs <new-owner> [--dry-run]
// Rewrites https://<owner>.github.io/mhs-billing-excel-addin/, github.com/<owner>/mhs-billing-excel-addin and
// "<owner>/mhs-billing-excel-addin" in the manifest and docs, then bumps the manifest <Version> so Office picks it up.
import fs from 'node:fs';

const [owner, flag] = process.argv.slice(2);
const dry = flag === '--dry-run';
if (!owner || !/^[A-Za-z0-9-]+$/.test(owner)) {
  console.error('Usage: node scripts/set-host.mjs <new-github-owner> [--dry-run]');
  process.exit(1);
}

const REPO = 'mhs-billing-excel-addin';
const manifest = fs.readFileSync('manifest.xml', 'utf8');
const cur = manifest.match(/https:\/\/([A-Za-z0-9-]+)\.github\.io\/mhs-billing-excel-addin\//);
if (!cur) { console.error('Could not find the current GitHub Pages URL in manifest.xml'); process.exit(1); }
const old = cur[1];
if (old.toLowerCase() === owner.toLowerCase()) { console.log(`Already pointed at ${owner}.`); process.exit(0); }

const files = ['manifest.xml', 'README.md', 'AUTOMATION.md', 'CLAUDE.md', 'docs/BUILD_SPEC.md', 'docs/POWER_PAGES.md', 'docs/SSO_SETUP.md']
  .filter((f) => fs.existsSync(f));
const esc = old.replace(/[-]/g, '\\-');
const swaps = [
  [new RegExp(`https://${esc}\\.github\\.io(?=[/<"\\s]|$)`, 'gim'), `https://${owner.toLowerCase()}.github.io`],
  [new RegExp(`api://${esc}\\.github\\.io/`, 'gi'), `api://${owner.toLowerCase()}.github.io/`], // SSO Application ID URI (also update it in Entra)
  [new RegExp(`github\\.com/${esc}/${REPO}`, 'gi'), `github.com/${owner}/${REPO}`],
  [new RegExp(`\\b${esc}/${REPO}`, 'g'), `${owner}/${REPO}`],
];

for (const f of files) {
  let text = fs.readFileSync(f, 'utf8');
  const before = text;
  for (const [re, to] of swaps) text = text.replace(re, to);
  if (f === 'manifest.xml') {
    text = text.replace(/<Version>(\d+)\.(\d+)\.(\d+)\.(\d+)<\/Version>/, (_, a, b) => `<Version>${a}.${Number(b) + 1}.0.0</Version>`);
  }
  if (text !== before) {
    console.log(`${dry ? 'would update' : 'updated'} ${f}`);
    if (!dry) fs.writeFileSync(f, text);
  }
}
console.log(`\n${old} -> ${owner}. New manifest: https://${owner.toLowerCase()}.github.io/${REPO}/manifest.xml`);
console.log('Next: npm run build && npm test, commit, push, then everyone re-uploads manifest.xml (or IT updates Integrated apps).');
