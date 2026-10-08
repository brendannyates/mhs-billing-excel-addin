// Build without dependencies (Node 22.6+): src/rules.ts -> rules.js (browser global BE),
// src/office/*.ts -> office-scripts/*.ts (paste-ready; rules appended where used), Power Pages bundle -> dist-pages/.
import { stripTypeScriptTypes } from "node:module";
import fs from "node:fs";
import path from "node:path";
const root = path.resolve(path.dirname(new URL(import.meta.url).pathname), "..");
const read = (p) => fs.readFileSync(path.join(root, p), "utf8");
const rules = read("src/rules.ts");
const names = [...rules.matchAll(/^(?:const|let|function)\s+([A-Za-z_]\w*)/gm)].map((m) => m[1]);
const js = stripTypeScriptTypes(rules);
fs.writeFileSync(path.join(root, "rules.js"),
  "/* Generated from src/rules.ts by scripts/build.mjs. Do not edit. */\n(function(){\n" + js +
  "\nvar api={" + names.map((n) => n + ":" + n).join(",") + "};\nglobalThis.BE=api;\n})();\n");
const USES_RULES = ["Automation.ts", "Digests.ts", "ClinicRecap.ts", "DigestsDv.ts"];
fs.mkdirSync(path.join(root, "office-scripts"), { recursive: true });
for (const f of fs.readdirSync(path.join(root, "src/office")).filter((f) => f.endsWith(".ts"))) {
  let out = "// " + f + " — paste into Excel > Automate > New script. Generated from src/office/" + f + (USES_RULES.includes(f) ? " + src/rules.ts" : "") + ". Do not edit here.\n\n" + read("src/office/" + f);
  if (USES_RULES.includes(f)) out += "\n\n// ======================= shared rules (src/rules.ts) =======================\n" + rules;
  fs.writeFileSync(path.join(root, "office-scripts", f), out);
}
// Power Pages edition (Dataverse): flat static site for `pac pages upload-code-site --compiledPath ./dist-pages`
const dp = path.join(root, "dist-pages");
fs.rmSync(dp, { recursive: true, force: true });
fs.mkdirSync(path.join(dp, "assets"), { recursive: true });
for (const f of ["rules.js", "app.js", "demo.js", "style.css"]) fs.copyFileSync(path.join(root, f), path.join(dp, f));
for (const f of ["dataverse.js", "config.js"]) fs.copyFileSync(path.join(root, "pages", f), path.join(dp, f));
// Same page as the add-in, minus Office: Dataverse source instead of Excel, full-width layout.
const pageHtml = read("index.html")
  .replace(/\s*<script src="https:\/\/appsforoffice[^>]*><\/script>/, "")
  .replace('<body>', '<body class="web">')
  .replace(/<script type="module" src="excel\.js[^"]*"><\/script>/, '<script src="config.js"></script>\n  <script src="dataverse.js?v=5"></script>');
if (pageHtml.includes("excel.js") || pageHtml.includes("appsforoffice")) throw new Error("pages index still references Office");
fs.writeFileSync(path.join(dp, "index.html"), pageHtml);
for (const f of fs.readdirSync(path.join(root, "assets"))) fs.copyFileSync(path.join(root, "assets", f), path.join(dp, "assets", f));
console.log("built rules.js (" + names.length + " exports), office-scripts/, dist-pages/");
