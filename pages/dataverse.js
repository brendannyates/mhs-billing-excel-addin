/* dataverse.js — Power Pages edition data source (tickets in Dataverse). Same interface as ExcelSource / DemoSource.
   Reads/writes through the Power Pages Web API (/_api/...) as the signed-in user, so table
   permissions on the site decide who can see and change what. */
(function () {
  "use strict";
  const BE = window.BE;
  const CFG = (window.BE_DATAVERSE && window.BE_DATAVERSE.sets) || {};
  const SET = {
    ticket: CFG.mhs_ticket || "mhs_tickets",
    activity: CFG.mhs_ticketactivity || "mhs_ticketactivities",
    setting: CFG.mhs_setting || "mhs_settings",
    holiday: CFG.mhs_holiday || "mhs_holidays",
    ref: CFG.mhs_refoption || "mhs_refoptions",
  };
  const offset = (ms) => -new Date(ms).getTimezoneOffset();
  const nowSerial = () => window.beNowSerial();
  const guidById = new Map();

  // ---- request token: Power Pages shell helper, falling back to the token endpoint ----
  async function token() {
    try {
      const sh = window.shell || (window.top && window.top.shell);
      if (sh && sh.getTokenDeferred) {
        return await new Promise((res, rej) => { const d = sh.getTokenDeferred(); d.done ? d.done(res).fail(rej) : d.then(res, rej); });
      }
    } catch (e) { /* fall through */ }
    const html = await (await fetch("/_layout/tokenhtml", { credentials: "same-origin" })).text();
    const m = html.match(/value="([^"]+)"/);
    if (!m) throw new Error("Couldn't get a Power Pages request token. Are you signed in?");
    return m[1];
  }
  async function req(method, url, body) {
    const headers = { Accept: "application/json", "Content-Type": "application/json", "X-Requested-With": "XMLHttpRequest", "OData-MaxVersion": "4.0", "OData-Version": "4.0" };
    if (method !== "GET") headers.__RequestVerificationToken = await token();
    else headers.Prefer = "odata.maxpagesize=5000";
    const res = await fetch(url.startsWith("http") || url.startsWith("/") ? url : "/_api/" + url, { method, credentials: "same-origin", headers, body: body ? JSON.stringify(body) : undefined });
    const text = await res.text();
    if (!res.ok) {
      let msg = text;
      try { const j = JSON.parse(text); msg = (j.error && (j.error.message || j.error.innererror && j.error.innererror.message)) || text; } catch (e) { /* raw */ }
      if (res.status === 403 || res.status === 401) msg = "No permission (" + res.status + "). Check the site's table permissions and Webapi site settings. " + msg;
      throw new Error(msg.slice(0, 400));
    }
    return text ? JSON.parse(text) : {};
  }
  async function getAll(url) {
    let out = [], next = url;
    while (next) { const r = await req("GET", next); out = out.concat(r.value || []); next = r["@odata.nextLink"] || null; }
    return out;
  }
  const TICKET_SELECT = ["mhs_ticketid"].concat(BE.DV_TICKET_FIELDS.map((f) => f[1])).join(",");

  const DataverseSource = {
    kind: "dataverse",

    async status() {
      await req("GET", SET.ticket + "?$select=mhs_formid&$top=1"); // throws a clear error if Web API / permissions aren't set up
      return { missingCols: [], master: true, raw: true, ref: true, activity: true, settings: true, holidays: true };
    },
    async setup() { return []; },
    async addColumns(c) { return c; },

    async load() {
      const [tickets, acts, settings, holidays, refs] = await Promise.all([
        getAll(SET.ticket + "?$select=" + TICKET_SELECT),
        getAll(SET.activity + "?$select=mhs_at,mhs_ticketnumber,mhs_name,mhs_oldvalue,mhs_newvalue,mhs_by,mhs_source&$orderby=mhs_at desc").catch(() => []),
        getAll(SET.setting + "?$select=mhs_name,mhs_value").catch(() => []),
        getAll(SET.holiday + "?$select=mhs_name,mhs_date").catch(() => []),
        getAll(SET.ref + "?$select=mhs_name,mhs_category,mhs_email,mhs_sort").catch(() => []),
      ]);
      const t = BE.dvTicketsToTable(tickets, offset);
      guidById.clear();
      t.rows.forEach((r, i) => { const id = BE.toNum(r[0]); if (id !== null) guidById.set(id, t.ids[i]); });
      // REF options -> the same column-per-category shape the REF sheet had
      const cats = new Map();
      refs.sort((a, b) => (a.mhs_sort || 0) - (b.mhs_sort || 0)).forEach((r) => {
        const c = r.mhs_category || "Other";
        if (!cats.has(c)) cats.set(c, []);
        cats.get(c).push(r);
      });
      const clin = cats.get("ClinicName") || [];
      if (clin.length) cats.set("ClinicEmail", clin.map((r) => ({ mhs_name: r.mhs_email || "" })));
      if (clin.length && !cats.has("Patient Primary Clinic:")) cats.set("Patient Primary Clinic:", clin);
      const headers = Array.from(cats.keys());
      const depth = Math.max(0, ...Array.from(cats.values()).map((v) => v.length));
      const rows = [];
      for (let i = 0; i < depth; i++) rows.push(headers.map((h) => (cats.get(h)[i] ? cats.get(h)[i].mhs_name : "")));
      const activity = BE.dvActivityRows(acts, offset).map((a) => [a.at, a.id, a.field, a.oldVal, a.newVal, a.by, a.source]);
      return {
        master: { headers: t.headers, rows: t.rows },
        ref: headers.length ? { headers, rows } : null,
        holidays: holidays.map((h) => h.mhs_date),
        settings: settings.map((s) => [s.mhs_name, s.mhs_value]),
        activity,
      };
    },

    /** Tickets arrive through the intake flow, so "sync" is just a reload. */
    async sync() { return { appended: [], updated: 0, baselined: 0, duplicates: [], at: nowSerial() }; },

    async saveTicket(id, edits, by) {
      const guid = guidById.get(id);
      if (!guid) throw new Error("Ticket #" + id + " not found. Refresh and try again.");
      const fields = edits.map((e) => BE.dvField(e.col)).filter(Boolean);
      const cur = await req("GET", SET.ticket + "(" + guid + ")?$select=" + fields.map((f) => f[1]).join(","));
      const patch = {};
      const act = [];
      const now = nowSerial();
      for (const e of edits) {
        const f = BE.dvField(e.col);
        if (!f) continue;
        const old = BE.dvToCell(f[2], cur[f[1]], offset);
        const same = f[2] === "bool" ? !!old === !!e.value : BE.cellEq(old, e.value);
        if (same) continue;
        patch[f[1]] = BE.cellToDv(f[2], e.value, offset);
        const show = (v) => (BE.DATE_COLS.indexOf(e.col) >= 0 ? (BE.toSerial(v) === null ? "" : BE.fmtDateTime(BE.toSerial(v))) : BE.norm(v));
        act.push({ mhs_name: e.col.replace(/:$/, ""), mhs_ticketnumber: id, mhs_oldvalue: show(old), mhs_newvalue: show(e.value), mhs_by: by, mhs_source: "Web", mhs_at: BE.serialToUtcIso(now, offset) });
      }
      if (!act.length) return { changed: 0 };
      patch.mhs_lastupdatedby = by;
      patch.mhs_lastupdatedat = BE.serialToUtcIso(now, offset);
      await req("PATCH", SET.ticket + "(" + guid + ")", patch);
      for (const a of act) { try { await req("POST", SET.activity, a); } catch (e) { console.warn("activity log failed", e); } }
      return { changed: act.length };
    },

    /** Refresh every 60s while the tab is visible so edits by others show up. */
    async watch(cb) {
      setInterval(() => { if (document.visibilityState === "visible") cb(); }, 60000);
    },
    openUrl(url) { window.open(url, "_blank", "noopener"); },

    /** Signed-in Power Pages user, if the platform exposes it. */
    currentUser() {
      try {
        const u = window.Microsoft && window.Microsoft.Dynamic365 && window.Microsoft.Dynamic365.Portal && window.Microsoft.Dynamic365.Portal.User;
        if (u && (u.email || u.userName)) return { email: String(u.email || u.userName).toLowerCase(), name: [u.firstName, u.lastName].filter(Boolean).join(" ") };
      } catch (e) { /* none */ }
      return null;
    },
  };
  window.DataverseSource = DataverseSource;
  if (!window.beNowSerial) window.beNowSerial = () => { const d = new Date(); return (d.getTime() - d.getTimezoneOffset() * 60000) / 86400000 + 25569; };
})();
