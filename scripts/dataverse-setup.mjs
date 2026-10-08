#!/usr/bin/env node
// MHS Billing Tickets (Power Pages edition) — one-time Dataverse setup + migration.
// Creates (idempotently) the publisher "mhs", solution "BillingEscalations" and 5 tables,
// seeds Settings / Holidays / REF options, imports existing tickets, and writes pages/config.js.
//
//   npm install
//   node scripts/dataverse-setup.mjs --env https://<org>.crm.dynamics.com [--data migration-data.json] [--dry-run]
//
// Sign-in uses a device code (you open a URL and paste a code). Needs System Customizer / System Admin
// in that environment. If your tenant blocks the default public client, pass --clientId <your app id>.
import fs from "node:fs";
import path from "node:path";
import vm from "node:vm";
const root = path.resolve(path.dirname(new URL(import.meta.url).pathname), "..");
const sandbox = vm.createContext({});
vm.runInContext(fs.readFileSync(path.join(root, "rules.js"), "utf8"), sandbox);
const BE = sandbox.BE;

const args = Object.fromEntries(process.argv.slice(2).reduce((a, v, i, arr) => {
  if (v.startsWith("--")) a.push([v.slice(2), arr[i + 1] && !arr[i + 1].startsWith("--") ? arr[i + 1] : true]);
  return a;
}, []));
if (!args.env) { console.error("Usage: node scripts/dataverse-setup.mjs --env https://<org>.crm.dynamics.com [--data migration-data.json] [--dry-run]"); process.exit(1); }
const ENV = String(args.env).replace(/\/$/, "");
const API = args.api ? String(args.api) : ENV + "/api/data/v9.2/"; // --api only for local testing
const DRY = !!args["dry-run"];
const SOLUTION = "BillingEscalations";
const PUBLISHER = { uniquename: "mhsoperations", friendlyname: "MHS Operations", customizationprefix: "mhs", customizationoptionvalueprefix: 72714 };

// ------------------------------------------------------------------ auth
async function getToken() {
  const msal = await import("@azure/msal-node");
  const app = new msal.PublicClientApplication({
    auth: { clientId: args.clientId || "51f81489-12ee-4a9e-aaae-a2591f45987d", authority: "https://login.microsoftonline.com/" + (args.tenant || "organizations") },
  });
  const r = await app.acquireTokenByDeviceCode({ scopes: [ENV + "/user_impersonation"], deviceCodeCallback: (d) => console.log("\n" + d.message + "\n") });
  return r.accessToken;
}
let TOKEN = "";
async function api(method, url, body, headers = {}) {
  if (DRY && method !== "GET") { console.log("  [dry-run]", method, url); return {}; }
  const res = await fetch(url.startsWith("http") ? url : API + url, {
    method,
    headers: { Authorization: "Bearer " + TOKEN, Accept: "application/json", "Content-Type": "application/json; charset=utf-8", "OData-Version": "4.0", "OData-MaxVersion": "4.0", ...headers },
    body: body ? JSON.stringify(body) : undefined,
  });
  if (res.status === 404) return null;
  const text = await res.text();
  if (!res.ok) throw new Error(method + " " + url + " → " + res.status + " " + text.slice(0, 600));
  return text ? JSON.parse(text) : {};
}
async function getAll(url) {
  let out = [], next = url;
  while (next) { const r = await api("GET", next, null, { Prefer: "odata.maxpagesize=5000" }); out = out.concat(r.value || []); next = r["@odata.nextLink"]; }
  return out;
}

// ------------------------------------------------------------------ metadata helpers
const L = (s) => ({ "@odata.type": "Microsoft.Dynamics.CRM.Label", LocalizedLabels: [{ "@odata.type": "Microsoft.Dynamics.CRM.LocalizedLabel", Label: s, LanguageCode: 1033 }] });
const REQ = { Value: "None", CanBeChanged: true, ManagedPropertyLogicalName: "canmodifyrequirementlevelsettings" };
const schemaOf = (logical) => "mhs_" + logical.slice(4).replace(/^./, (c) => c.toUpperCase());
function attr(logical, type, label) {
  const base = { SchemaName: schemaOf(logical), DisplayName: L(label), Description: L(label), RequiredLevel: REQ };
  switch (type) {
    case "text": return { "@odata.type": "Microsoft.Dynamics.CRM.StringAttributeMetadata", ...base, MaxLength: 400, FormatName: { Value: "Text" } };
    case "memo": return { "@odata.type": "Microsoft.Dynamics.CRM.MemoAttributeMetadata", ...base, MaxLength: 10000, Format: "TextArea" };
    case "int": return { "@odata.type": "Microsoft.Dynamics.CRM.IntegerAttributeMetadata", ...base, Format: "None", MinValue: -2147483648, MaxValue: 2147483647 };
    case "num": return { "@odata.type": "Microsoft.Dynamics.CRM.DecimalAttributeMetadata", ...base, Precision: 2, MinValue: -100000000000, MaxValue: 100000000000 };
    case "dt": return { "@odata.type": "Microsoft.Dynamics.CRM.DateTimeAttributeMetadata", ...base, Format: "DateAndTime", DateTimeBehavior: { Value: "UserLocal" } };
    case "date": return { "@odata.type": "Microsoft.Dynamics.CRM.DateTimeAttributeMetadata", ...base, Format: "DateOnly", DateTimeBehavior: { Value: "DateOnly" } };
    case "bool": return { "@odata.type": "Microsoft.Dynamics.CRM.BooleanAttributeMetadata", ...base, DefaultValue: false,
      OptionSet: { "@odata.type": "Microsoft.Dynamics.CRM.BooleanOptionSetMetadata", TrueOption: { Value: 1, Label: L("Yes") }, FalseOption: { Value: 0, Label: L("No") }, OptionSetType: "Boolean" } };
    default: throw new Error("unknown type " + type);
  }
}
const TABLES = [
  { logical: "mhs_ticket", set: "mhs_tickets", name: "Billing Ticket", plural: "Billing Tickets", primary: "Title",
    cols: BE.DV_TICKET_FIELDS.map((f) => [f[1], f[2], f[3]]), key: "mhs_formid" },
  { logical: "mhs_ticketactivity", set: "mhs_ticketactivities", name: "Ticket Activity", plural: "Ticket Activity", primary: "Field",
    cols: BE.DV_ACTIVITY_FIELDS.filter((f) => f[1] !== "mhs_name").map((f) => [f[1], f[2], f[3]]) },
  { logical: "mhs_setting", set: "mhs_settings", name: "Escalation Setting", plural: "Escalation Settings", primary: "Setting",
    cols: [["mhs_value", "text", "Value"], ["mhs_notes", "text", "Notes"]] },
  { logical: "mhs_holiday", set: "mhs_holidays", name: "Holiday", plural: "Holidays", primary: "Holiday", cols: [["mhs_date", "date", "Date"]] },
  { logical: "mhs_refoption", set: "mhs_refoptions", name: "Reference Option", plural: "Reference Options", primary: "Value",
    cols: [["mhs_category", "text", "Category"], ["mhs_email", "text", "Email"], ["mhs_sort", "int", "Sort order"]] },
];

async function ensurePublisherAndSolution() {
  let pub = (await api("GET", `publishers?$select=publisherid&$filter=uniquename eq '${PUBLISHER.uniquename}'`)).value[0];
  if (!pub) { console.log("• Creating publisher 'mhs'"); await api("POST", "publishers", PUBLISHER); pub = (await api("GET", `publishers?$select=publisherid&$filter=uniquename eq '${PUBLISHER.uniquename}'`)).value[0]; }
  const sol = (await api("GET", `solutions?$select=solutionid&$filter=uniquename eq '${SOLUTION}'`)).value[0];
  if (!sol) {
    console.log("• Creating solution 'BillingEscalations'");
    await api("POST", "solutions", { uniquename: SOLUTION, friendlyname: "Billing Escalations", version: "1.0.0.0", "publisherid@odata.bind": "/publishers(" + (pub ? pub.publisherid : "00000000-0000-0000-0000-000000000000") + ")" });
  }
}
async function ensureTable(t) {
  const sol = { "MSCRM.SolutionUniqueName": SOLUTION };
  let def = await api("GET", `EntityDefinitions(LogicalName='${t.logical}')?$select=EntitySetName`);
  let created = false;
  if (!def) {
    console.log("• Creating table " + t.name);
    await api("POST", "EntityDefinitions", {
      "@odata.type": "Microsoft.Dynamics.CRM.EntityMetadata", SchemaName: schemaOf(t.logical), EntitySetName: t.set,
      DisplayName: L(t.name), DisplayCollectionName: L(t.plural), Description: L("Billing Escalations — " + t.plural),
      OwnershipType: "OrganizationOwned", HasActivities: false, HasNotes: false, IsActivity: false,
      Attributes: [{ "@odata.type": "Microsoft.Dynamics.CRM.StringAttributeMetadata", SchemaName: "mhs_Name", IsPrimaryName: true, MaxLength: 400,
        FormatName: { Value: "Text" }, RequiredLevel: REQ, DisplayName: L(t.primary), Description: L(t.primary) }],
    }, sol);
    created = true;
    def = DRY ? { EntitySetName: t.set } : await api("GET", `EntityDefinitions(LogicalName='${t.logical}')?$select=EntitySetName`);
  }
  const existing = new Set();
  if (!(DRY && created)) {
    const r = await api("GET", `EntityDefinitions(LogicalName='${t.logical}')/Attributes?$select=LogicalName`);
    ((r && r.value) || []).forEach((a) => existing.add(a.LogicalName));
  }
  for (const [logical, type, label] of t.cols) {
    if (existing.has(logical)) continue;
    console.log("  + column " + label + " (" + logical + ")");
    await api("POST", `EntityDefinitions(LogicalName='${t.logical}')/Attributes`, attr(logical, type, label), sol);
  }
  if (t.key) {
    const keys = DRY && created ? [] : ((await api("GET", `EntityDefinitions(LogicalName='${t.logical}')/Keys?$select=SchemaName`)) || { value: [] }).value;
    if (!keys.some((k) => k.SchemaName === "mhs_FormIdKey")) {
      console.log("  + alternate key on Ticket #");
      await api("POST", `EntityDefinitions(LogicalName='${t.logical}')/Keys`, { SchemaName: "mhs_FormIdKey", DisplayName: L("Ticket # key"), KeyAttributes: [t.key] }, sol);
    }
  }
  return def.EntitySetName || t.set;
}

// ------------------------------------------------------------------ seed + migrate
const offset = (ms) => BE.pacificOffsetMin(ms);
async function seed(sets, data) {
  const settings = await getAll(sets.mhs_setting + "?$select=mhs_name");
  if (!settings.length) {
    console.log("• Seeding Settings");
    for (const [k, v, n] of (data && data.settings) || []) await api("POST", sets.mhs_setting, { mhs_name: k, mhs_value: String(v), mhs_notes: n || "" });
  }
  const hol = await getAll(sets.mhs_holiday + "?$select=mhs_name");
  if (!hol.length) {
    console.log("• Seeding Holidays");
    for (const [d, n] of (data && data.holidays) || []) await api("POST", sets.mhs_holiday, { mhs_name: n, mhs_date: d });
  }
  const ref = await getAll(sets.mhs_refoption + "?$select=mhs_name");
  if (!ref.length && data && data.ref) {
    console.log("• Seeding REF options (" + data.ref.length + ")");
    for (const r of data.ref) await api("POST", sets.mhs_refoption, { mhs_name: r.value, mhs_category: r.category, mhs_email: r.email || null, mhs_sort: r.sort });
  }
}
async function migrate(sets, data) {
  if (!data || !data.tickets) { console.log("• No --data file: skipping ticket import"); return; }
  const have = new Set((await getAll(sets.mhs_ticket + "?$select=mhs_formid")).map((r) => r.mhs_formid));
  const idx = BE.buildIndex(data.tickets.headers);
  let n = 0;
  for (const row of data.tickets.rows) {
    const id = BE.toNum(BE.cell(row, idx, BE.COL.id));
    if (id === null || have.has(id)) continue;
    const rec = { mhs_name: "#" + id + " " + BE.norm(BE.cell(row, idx, BE.COL.patient)) };
    for (const [col, logical, type] of BE.DV_TICKET_FIELDS) {
      const c = BE.ci(idx, col);
      if (c < 0) continue;
      const v = BE.cellToDv(type, row[c], offset);
      if (v !== null) rec[logical] = v;
    }
    await api("POST", sets.mhs_ticket, rec);
    n++;
    if (n % 10 === 0) console.log("  … " + n + " tickets");
  }
  console.log("• Imported " + n + " tickets (" + have.size + " already present)");
}

// ------------------------------------------------------------------ main
(async () => {
  const defaults = JSON.parse(fs.readFileSync(path.join(root, "pages/seed-defaults.json"), "utf8"));
  const data = Object.assign({}, defaults, args.data ? JSON.parse(fs.readFileSync(String(args.data), "utf8")) : {});
  TOKEN = args.token ? String(args.token) : await getToken(); // --token only for local testing
  const who = await api("GET", "WhoAmI");
  console.log("Signed in. User id " + (who && who.UserId) + " on " + ENV + (DRY ? "  (dry run)" : ""));
  await ensurePublisherAndSolution();
  const sets = {};
  for (const t of TABLES) sets[t.logical] = await ensureTable(t);
  console.log("• Publishing customizations…");
  await api("POST", "PublishAllXml", {});
  await seed(sets, data);
  await migrate(sets, data);
  const cfg = "/* Generated by scripts/dataverse-setup.mjs */\nwindow.BE_DATAVERSE = " + JSON.stringify({ sets }, null, 2) + ";\n";
  fs.writeFileSync(path.join(root, "pages/config.js"), cfg);
  console.log("\nDone. Entity sets:", sets);
  console.log("Next: npm run build, then upload dist-pages/ (see docs/POWER_PAGES.md).");
})().catch((e) => { console.error("\nSetup failed:", e.message); process.exit(1); });
