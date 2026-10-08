// Turn on Microsoft 365 single sign-on in the add-in manifest once the Entra app registration exists.
// Usage: node scripts/set-sso.mjs <application-client-id> [--dry-run]
// Adds (or replaces) <WebApplicationInfo> inside <VersionOverrides>, with the Application ID URI
// api://<pages-host>/<client-id>, and bumps the manifest <Version>. See docs/SSO_SETUP.md.
import fs from 'node:fs';

const [id, flag] = process.argv.slice(2);
if (!id || !/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(id)) {
  console.error('Usage: node scripts/set-sso.mjs <application-client-id (GUID)> [--dry-run]');
  process.exit(1);
}
let xml = fs.readFileSync('manifest.xml', 'utf8');
const host = (xml.match(/<SourceLocation DefaultValue="https:\/\/([^/"]+)\//) || [])[1];
if (!host) { console.error('Could not read the host from <SourceLocation> in manifest.xml'); process.exit(1); }
const block =
  '    <WebApplicationInfo>\n' +
  `      <Id>${id}</Id>\n` +
  `      <Resource>api://${host}/${id}</Resource>\n` +
  '      <Scopes>\n        <Scope>openid</Scope>\n        <Scope>profile</Scope>\n      </Scopes>\n' +
  '    </WebApplicationInfo>\n';
xml = xml.replace(/\s*<WebApplicationInfo>[\s\S]*?<\/WebApplicationInfo>\n?/, '\n');
xml = xml.replace(/(\n\s*<\/Resources>\n)/, `$1${block}`);
xml = xml.replace(/<Version>(\d+)\.(\d+)\.(\d+)\.(\d+)<\/Version>/, (_, a, b, c) => `<Version>${a}.${b}.${Number(c) + 1}.0</Version>`);
if (flag === '--dry-run') { console.log(xml.match(/<Version>.*<\/Version>/)[0] + '\n' + block); process.exit(0); }
fs.writeFileSync('manifest.xml', xml);
console.log(`SSO enabled for ${id} (Application ID URI api://${host}/${id}).`);
console.log('Next: npm run build && npm test, commit, push, then everyone re-uploads manifest.xml (or IT updates Integrated apps).');
