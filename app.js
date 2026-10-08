/* MHS Billing Tickets — one UI for the Excel task pane, its full-screen window and the Power Pages site.
   Vanilla JS; state -> render() -> event delegation. Rules come from rules.js (window.BE). */
(function () {
  "use strict";
  const BE = window.BE;
  const C = BE.COL;
  const VERSION = "2.0.0";
  const MODE = new URLSearchParams(location.search).get("mode") || "pane";
  const DEFAULT_FORM = "https://forms.cloud.microsoft/Pages/ResponsePage.aspx?id=8Sq2y8CkOEWdeGKInLNdYw1_rrxVxaFImLkzEWnt0GFUQUtJVFg2VzU5UEowM01JMUgzUUpKMTZPVi4u";
  const REF_MAP = {
    [C.dept]: "Department Submitting Ticket:", [C.clinic]: "Patient Primary Clinic:", [C.urgency]: "Urgency Level:",
    [C.source]: "Source of Inquiry", [C.taskType]: "Task Type", [C.status]: "Status",
    [C.falseVerif]: "Was false verification of balance/charge given by Arietis?",
    [C.serviceRecovery]: "Patient Feedback/Service Recovery Flag:", [C.errorSource]: "Source of Error:", [C.outcome]: "Outcome:",
  };
  const FALLBACK = {
    [C.status]: Object.values(BE.STATUS), [C.urgency]: ["Normal", "High", "Critical"], [C.falseVerif]: ["Yes", "No"],
    [C.serviceRecovery]: ["Yes", "No"], [C.errorSource]: ["Arietis", "MHS", "Arietis + MHS"],
    [C.outcome]: ["Refund", "Rebill/Recode Claims", "Patient Plan Education", "BIDF Incorrect - Needed Correction"],
  };

  const store = {
    get(k, d) { try { const v = localStorage.getItem("be." + k); return v === null ? d : v; } catch (e) { return d; } },
    set(k, v) { try { localStorage.setItem("be." + k, v); } catch (e) { /* private mode */ } },
  };

  const PROFILE_KEY = "mhsBillingProfile";
  let profile = { name: "", email: "", role: "", clinics: [] };
  try { profile = { ...profile, ...JSON.parse(localStorage.getItem(PROFILE_KEY) || "{}") }; } catch (e) { /* private mode */ }
  const S = {
    src: null, ready: false, me: BE.normKey(profile.email), tab: store.get("tab", "mine"), detail: null,
    tickets: [], byId: new Map(), settings: BE.readSettings([]), ctx: null, refCols: new Map(), clinics: [],
    clinicEmails: new Map(), people: new Map(), activity: [], lastSync: null, setupNeeded: false,
    f: { scope: "assigned", recapMode: "am", recapClinic: "", actq: "", ops: "queue", mine: "open", clinic: store.get("clinic", ""), stage: "open", actWho: "all", actClinic: "", range: "all", q: "", find: "" },
  };

  // ------------------------------------------------------------------ utils
  const $ = (s, r) => (r || document).querySelector(s);
  const esc = (v) => String(v === null || v === undefined ? "" : v).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const nowSerial = () => window.beNowSerial();
  const shortClinic = (c) => c.replace(/ - (CA|WA|TX)$/, "");
  const pctTxt = (a, b) => { const p = BE.pct(a, b); return p === null ? "—" : p + "%"; };
  const avg = (s, n) => (n ? (s / n).toFixed(1) : "—");
  const isoDay = (s) => { if (s === null || s === undefined) return ""; const d = BE.serialToDate(Math.floor(s)); return d.getUTCFullYear() + "-" + String(d.getUTCMonth() + 1).padStart(2, "0") + "-" + String(d.getUTCDate()).padStart(2, "0"); };
  let toastTimer = null;
  function toast(msg, err) {
    const t = $("#toast");
    t.textContent = msg; t.className = "toast" + (err ? " err" : ""); t.hidden = false;
    clearTimeout(toastTimer); toastTimer = setTimeout(() => (t.hidden = true), err ? 7000 : 3500);
  }
  function busy(on, label) {
    const s = $("#syncState");
    s.classList.toggle("busy", !!on);
    s.textContent = on ? label || "Working…" : S.lastSync ? "Synced " + BE.fmtTime(S.lastSync) : "";
    $("#app").setAttribute("aria-busy", on ? "true" : "false");
  }
  function meName() {
    if (profile.name) return profile.name;
    const p = S.people.get(S.me);
    return p ? p : S.meNameHint ? S.meNameHint : S.me ? S.me.split("@")[0] : "";
  }
  /** Assigned-clinic scope (profile). Empty profile clinics = all clinics. */
  function inScope(t) { return S.f.scope === "all" || !profile.clinics.length || profile.clinics.indexOf(t.clinic) >= 0; }
  function scopeSel() {
    if (!profile.clinics.length) return "";
    return `<select data-filter="scope" aria-label="Clinic scope"><option value="assigned" ${S.f.scope === "assigned" ? "selected" : ""}>My clinics (${profile.clinics.length})</option><option value="all" ${S.f.scope === "all" ? "selected" : ""}>All clinics</option></select>`;
  }

  // ------------------------------------------------------------------ full screen (Portal): dialog proxies every call to the task pane
  // The dialog window can't call Excel directly, so it asks the pane (which owns the workbook session) over
  // Office dialog messaging. Messages are chunked because each one is size-limited.
  const CHUNK = 24000;
  function sendChunks(sendFn, obj) {
    const str = JSON.stringify(obj), k = Math.random().toString(36).slice(2), n = Math.max(1, Math.ceil(str.length / CHUNK));
    for (let i = 0; i < n; i++) sendFn(JSON.stringify({ k, i, n, d: str.slice(i * CHUNK, (i + 1) * CHUNK) }));
  }
  function receiver(cb) {
    const parts = new Map();
    return (msg) => {
      let m; try { m = JSON.parse(msg); } catch (e) { return; }
      if (!m || m.k === undefined) return;
      if (m.k === "close") { cb({ close: true }); return; }
      const p = parts.get(m.k) || []; p[m.i] = m.d; parts.set(m.k, p);
      if (p.filter((x) => x !== undefined).length === m.n) { parts.delete(m.k); cb(JSON.parse(p.join(""))); }
    };
  }
  function canFullScreen() {
    try { return Office.context.requirements.isSetSupported("DialogApi", "1.2"); } catch (e) { return false; }
  }
  let dlg = null;
  function openFull() {
    if (dlg) return;
    const url = location.origin + location.pathname + "?mode=dialog";
    Office.context.ui.displayDialogAsync(url, { height: 94, width: 94, displayInIframe: false }, (r) => {
      if (r.status !== Office.AsyncResultStatus.Succeeded) { toast("Couldn't open full screen: " + (r.error && r.error.message), true); return; }
      dlg = r.value;
      const rx = receiver((m) => {
        if (m.close) { try { dlg.close(); } catch (e) { /* closed */ } dlg = null; refresh(false, true); return; }
        Promise.resolve().then(() => window.ExcelSource[m.method].apply(window.ExcelSource, m.args || []))
          .then((v) => ({ id: m.id, ok: true, v }), (e) => ({ id: m.id, ok: false, e: (e && e.message) || String(e) }))
          .then((res) => { try { sendChunks((x) => dlg.messageChild(x), res); } catch (e) { /* dialog closed */ } });
      });
      dlg.addEventHandler(Office.EventType.DialogMessageReceived, (ev) => rx(ev.message));
      dlg.addEventHandler(Office.EventType.DialogEventReceived, () => { dlg = null; refresh(false, true); });
    });
  }
  const pendingCalls = new Map();
  function dialogCall(method, ...args) {
    return new Promise((res, rej) => {
      const id = Math.random().toString(36).slice(2);
      pendingCalls.set(id, { res, rej });
      sendChunks((x) => Office.context.ui.messageParent(x), { id, method, args });
      setTimeout(() => { if (pendingCalls.has(id)) { pendingCalls.delete(id); rej(new Error("The task pane didn't answer. Keep the MHS Billing Tickets pane open behind this window.")); } }, 90000);
    });
  }
  const DialogSource = {
    kind: "excel-dialog",
    init() {
      const rx = receiver((m) => { const p = pendingCalls.get(m.id); if (!p) return; pendingCalls.delete(m.id); m.ok ? p.res(m.v) : p.rej(new Error(m.e)); });
      Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, (ev) => rx(ev.message));
    },
    status: () => dialogCall("status"), setup: (me) => dialogCall("setup", me), load: () => dialogCall("load"), sync: (u) => dialogCall("sync", u),
    saveTicket: (id, edits, by) => dialogCall("saveTicket", id, edits, by), setWorkbookUrl: (u) => dialogCall("setWorkbookUrl", u),
    addColumns: (c) => Promise.resolve(c),
    async watch(cb) { setInterval(() => { if (document.visibilityState === "visible" && !isEditing()) cb(); }, 30000); },
    openUrl(url) { window.open(url, "_blank", "noopener"); },
  };

  // ------------------------------------------------------------------ data
  async function boot() {
    const inExcel = MODE !== "dialog" && typeof Office !== "undefined" && Office.context && Office.context.host === Office.HostType.Excel;
    const onPages = !!window.DataverseSource && (!!window.BE_DATAVERSE || /\.(powerappsportals\.com|powerpages\.microsoft\.com)$/i.test(location.hostname)) && !/[?&]demo\b/.test(location.search);
    if (MODE === "dialog") DialogSource.init();
    S.src = MODE === "dialog" ? DialogSource : inExcel ? window.ExcelSource : onPages ? window.DataverseSource : window.DemoSource;
    if (S.src.kind === "demo") showBanner("Demo mode — synthetic data. Open from Excel (Home › MHS Billing Tickets) to work on the live workbook.");
    if (S.src.kind === "dataverse") $("#btnSync").title = "Refresh";
    if (S.src.currentUser) { const u = S.src.currentUser(); if (u && u.email) { S.me = u.email; if (u.name) S.meNameHint = u.name; } }
    if (inExcel && canFullScreen()) $("#btnFull").hidden = false;
    if (MODE === "dialog") { document.body.classList.add("web"); $("#btnFull").hidden = false; $("#btnFull").title = "Close full screen"; }
    try {
      if (S.src.kind === "excel" || S.src.kind === "excel-dialog") {
        busy(true, "Preparing workbook…");
        const added = await S.src.setup(S.me); // idempotent Initialize / refresh
        if (added.length) toast("Workbook prepared: " + added.length + " update" + (added.length > 1 ? "s" : "") + " (columns, Settings, Activity Log). Raw Data untouched.");
      }
      await refresh(true);
      S.src.watch(debounce(() => refresh(false, true), 1500));
      if (S.src.kind === "excel") setInterval(() => { if (!S._busy && document.visibilityState === "visible" && !isEditing()) refresh(true, true); }, 30000);
      if (!profile.email && S.src.kind !== "dataverse") pickMe();
    } catch (e) { renderFatal(e.message || String(e)); }
  }
  function isEditing() {
    const a = document.activeElement;
    return !!$("#overlay") || (S.detail !== null && detailEdits().length > 0) || S.review || (a && /INPUT|TEXTAREA|SELECT/.test(a.tagName) && a.id !== "find");
  }
  function debounce(fn, ms) { let t; return () => { clearTimeout(t); t = setTimeout(fn, ms); }; }

  async function refresh(doSync, quiet, force) {
    if (S._busy) return;
    S._busy = true;
    try {
      if (doSync) {
        busy(true, "Syncing…");
        const r = await S.src.sync(S.me);
        S.lastSync = r.at;
        if (r.appended.length) toast("Added " + r.appended.length + " new ticket" + (r.appended.length > 1 ? "s" : "") + ": #" + r.appended.join(", #"));
        else if (r.updated) toast("Merged " + r.updated + " form edit" + (r.updated > 1 ? "s" : "") + " into Master");
        if (r.duplicates && r.duplicates.length) showBanner("Duplicate ticket Ids on Master: #" + r.duplicates.join(", #") + " — delete the extra rows.");
      }
      busy(true, quiet ? "Updating…" : "Loading…");
      const d = await S.src.load();
      const dirty = quiet && !force && S.detail !== null && $("#editForm") && detailEdits().length > 0;
      ingest(d);
      S.ready = true;
      busy(false);
      if (dirty) { toast("Master changed in the background. Your unsaved edits are kept; save to apply."); return; }
      render();
    } catch (e) {
      busy(false);
      toast(e.message || String(e), true);
    } finally { S._busy = false; }
  }

  function ingest(d) {
    S.settings = BE.readSettings(d.settings || []);
    const hol = BE.holidaySet(d.holidays || []);
    S.ctx = { now: nowSerial(), hol, receiptDays: S.settings.receiptDays, resolutionDays: S.settings.resolutionDays, staleDays: S.settings.staleDays };
    S.tickets = BE.loadTickets(d.master.headers, d.master.rows, S.ctx);
    S.byId = new Map(S.tickets.map((t) => [t.id, t]));
    S.refCols = new Map();
    S.clinicEmails = new Map();
    if (d.ref) {
      d.ref.headers.forEach((h, i) => {
        const vals = [];
        d.ref.rows.forEach((r) => { const v = BE.norm(r[i]); if (v && v !== "[Free Text]" && vals.indexOf(v) < 0) vals.push(v); });
        S.refCols.set(BE.normKey(h), vals);
      });
      const ix = BE.buildIndex(d.ref.headers);
      d.ref.rows.forEach((r) => { const n = BE.norm(BE.cell(r, ix, "ClinicName")); const e = BE.norm(BE.cell(r, ix, "ClinicEmail")); if (n && e) S.clinicEmails.set(n.toLowerCase(), e); });
    }
    const clin = new Set(options(C.clinic));
    S.tickets.forEach((t) => clin.add(t.clinic));
    S.clinics = Array.from(clin).sort((a, b) => (a === "(No clinic)") - (b === "(No clinic)") || a.localeCompare(b));
    S.people = new Map();
    S.tickets.forEach((t) => { if (t.ownerEmail && !S.people.has(t.ownerEmail)) S.people.set(t.ownerEmail, t.ownerName || t.ownerEmail); });
    S.activity = (d.activity || []).map((r) => ({ at: BE.toSerial(r[0]) || 0, id: BE.toNum(r[1]) || 0, field: BE.norm(r[2]), oldVal: BE.norm(r[3]), newVal: BE.norm(r[4]), by: BE.norm(r[5]), source: BE.norm(r[6]) }));
    if (!S.f.clinic || S.clinics.indexOf(S.f.clinic) < 0) S.f.clinic = profile.clinics.find((c) => S.clinics.indexOf(c) >= 0) || mostCommonClinic();
  }
  function mostCommonClinic() {
    const c = new Map();
    S.tickets.filter((t) => !S.me || t.ownerEmail === S.me).forEach((t) => c.set(t.clinic, (c.get(t.clinic) || 0) + 1));
    let best = S.clinics[0] || "", n = -1;
    c.forEach((v, k) => { if (v > n) { best = k; n = v; } });
    return best;
  }
  function options(col) {
    const k = BE.normKey(REF_MAP[col] || "");
    const v = S.refCols.get(k);
    return v && v.length ? v : FALLBACK[col] || [];
  }

  // ------------------------------------------------------------------ shared renderers
  function stageClass(t) {
    return { "Awaiting receipt": "st-await", "With Arietis": "st-with", "Pending call": "st-call", Closed: "st-closed" }[t.stage] || "st-await";
  }
  function slaBadge(t) {
    if (!t.isOpen) {
      const r = t.resolution;
      if (r.state === "late") return `<span class="sla bad" title="Resolved after the ${S.ctx.resolutionDays}-BD SLA">Closed +${r.bdOver} BD</span>`;
      return `<span class="sla ok">Closed ${r.at !== null ? BE.fmtShort(r.at) : ""}</span>`;
    }
    const useRcpt = t.stage === "Awaiting receipt" && t.receipt.at === null;
    const s = useRcpt ? t.receipt : t.resolution;
    const k = useRcpt ? "Receipt" : "Resolve";
    if (s.state === "breached") return `<span class="sla bad" title="${k} SLA missed by ${s.bdOver} business days">${k} +${s.bdOver} BD</span>`;
    if (s.state === "due-today") return `<span class="sla warn">${k} due today</span>`;
    return `<span class="sla neutral" title="${k} due ${BE.fmtDate(s.due)}">${k} by ${BE.weekdayName(s.due)} ${BE.fmtShort(s.due)}</span>`;
  }
  function flagChips(t) {
    return t.flags.filter((f) => f !== "EHR task").map((f) => `<span class="flag ${f === "Critical" || f === "False verification" ? "red" : f === "High" || f === "Service recovery" ? "amber" : "gray"}">${esc(f === "Service recovery" ? "Svc recovery" : f === "False verification" ? "False verif." : f)}</span>`).join(" ");
  }
  const ISSUE_SHORT = [[/^balance dispute/i, "Balance dispute"], [/^duplicate/i, "Duplicate balance"], [/^refund/i, "Refund"], [/callback/i, "Callback request"], [/^superbill|eob/i, "Superbill / EOB"], [/unable to contact/i, "Can't reach Arietis"], [/not billing under ins/i, "Not billed to insurance"]];
  function shortIssue(x) { if (!x) return "Escalation"; for (const [re, l] of ISSUE_SHORT) if (re.test(x)) return l; return x.replace(/\s*\(.*\)$/, ""); }
  /** Collapsed ticket row: one line, like a ticket queue. Everything else lives in the detail view. */
  function row(t) {
    const issue = shortIssue(t.taskTypes[0]) + (t.taskTypes.length > 1 ? " +" + (t.taskTypes.length - 1) : "");
    const dot = t.actions.length ? `<span class="dot sev-${t.topSev}" title="${esc(t.actions[0].text)}"></span>` : `<span class="dot none"></span>`;
    const chip = { "Awaiting receipt": "Awaiting", "With Arietis": "With Arietis", "Pending call": "Pt call", Closed: "Closed" }[t.stage] || t.stage;
    return `<button class="trow" data-open="${t.id}" aria-label="Ticket ${t.id}, ${esc(issue)}, ${esc(t.stage)}${t.actions.length ? ", needs action" : ""}">${dot}<span class="tid">#${t.id}</span><span class="pid"><span class="pt">${esc(t.patient || "—")}</span><span class="mrn">${esc(t.mrn ? "MRN " + t.mrn : "no MRN")}</span></span><span class="subj">${esc(issue)}</span><span class="chip ${stageClass(t)}" title="${esc(t.status)}">${esc(chip)}</span></button>`;
  }
  function matches(t, q) {
    const k = q.toLowerCase().replace(/^#/, "").replace(/^mrn\s*/, "").trim();
    if (!k) return true;
    return t.mrn.indexOf(k) >= 0 || String(t.id) === k || t.patient.toLowerCase().indexOf(k) >= 0;
  }
  function list(ts, o, emptyMsg) {
    if (!ts.length && S.f.find && ["mine", "action", "ops", "clinic"].indexOf(S.tab) >= 0) return `<div class="empty">No match for “${esc(S.f.find)}” here. <button class="linkish" data-act="findAll">Search all tickets</button></div>`;
    if (!ts.length) return `<div class="empty">${esc(emptyMsg || "Nothing here.")}</div>`;
    const lim = (o && o.limit) || 400;
    return `<div class="list">${ts.slice(0, lim).map((t) => row(t, o)).join("")}</div>${ts.length > lim ? `<div class="muted" style="margin:6px 2px">Showing ${lim} of ${ts.length}. Narrow the filter to see more.</div>` : ""}`;
  }
  function sec(title, n, extra) {
    return `<div class="sec">${esc(title)}${n === null || n === undefined ? "" : ` <span class="n">${n}</span>`}<span class="sp"></span>${extra || ""}</div>`;
  }
  function seg(name, items, cur) {
    return `<div class="seg" role="group">${items.map(([v, l, c]) => `<button data-seg="${name}" data-v="${esc(v)}" aria-pressed="${v === cur}">${esc(l)}${c !== undefined ? `<span class="c">${c}</span>` : ""}</button>`).join("")}</div>`;
  }
  function kpi(label, value, cls, act) {
    const tag = act ? "button" : "div";
    return `<${tag} class="kpi ${cls || ""}" ${act || ""}><div class="v">${esc(value)}</div><div class="k">${esc(label)}</div></${tag}>`;
  }
  const attn = (ts) => BE.sortForAttention(ts);

  // ------------------------------------------------------------------ views
  function viewMine() {
    if (!S.me) return `<div class="setup"><h3>Who are you?</h3><p class="muted">My tickets shows escalations you submitted (by your Microsoft Forms email) and ones you're CC'd on.</p><button class="btn primary" data-act="pickMe">Choose my name</button></div>`;
    const pool = S.tickets.filter((t) => matches(t, S.f.find));
    const mine = pool.filter((t) => t.ownerEmail === S.me);
    const cc = pool.filter((t) => t.ownerEmail !== S.me && t.cc.indexOf(S.me) >= 0);
    const open = attn(mine.filter((t) => t.isOpen));
    const needs = attn(mine.filter((t) => t.actions.length));
    const closed = mine.filter((t) => !t.isOpen).sort((a, b) => (b.resolution.at || 0) - (a.resolution.at || 0));
    const past = open.filter((t) => t.topSev === 1).length;
    const mineAll = mine.concat(cc);
    const openAll = attn(mineAll.filter((t) => t.isOpen));
    const needAll = attn(mineAll.filter((t) => t.actions.length));
    const closedAll = mineAll.filter((t) => !t.isOpen).sort((a, b) => (b.resolution.at || 0) - (a.resolution.at || 0));
    if (["open", "action", "closed"].indexOf(S.f.mine) < 0) S.f.mine = "open";
    let h = `<div class="toolbar">${seg("mine", [["open", "Open", openAll.length], ["action", "Needs action", needAll.length], ["closed", "Closed", closedAll.length]], S.f.mine)}</div>`;
    if (S.f.mine === "open") h += list(openAll, {}, "No open tickets.");
    else if (S.f.mine === "action") h += list(needAll, {}, "You're all caught up.");
    else h += list(closedAll, {}, "No closed tickets yet.");
    return h;
  }

  function viewAction() {
    let ts = S.tickets.filter((t) => t.actions.length && inScope(t) && matches(t, S.f.find));
    if (S.f.actWho === "mine" && S.me) ts = ts.filter((t) => t.ownerEmail === S.me || t.cc.indexOf(S.me) >= 0);
    if (S.f.actClinic) ts = ts.filter((t) => t.clinic === S.f.actClinic);
    const vendor = ts.filter((t) => t.isOpen && t.actions.some((a) => a.who === "arietis"));
    const breached = vendor.filter((t) => t.topSev === 1);
    const crit = attn(ts.filter((t) => t.topSev === 1));
    const warn = attn(ts.filter((t) => t.topSev === 2));
    const info = attn(ts.filter((t) => t.topSev === 3));
    const owners = new Set(ts.filter((t) => t.actions.some((a) => a.who === "owner")).map((t) => t.ownerEmail));
    let h = `<div class="toolbar">${seg("actWho", [["all", "Everyone"], ["mine", "Mine"]], S.f.actWho)}${scopeSel()}
      <select data-filter="actClinic" aria-label="Clinic"><option value="">All clinics</option>${S.clinics.map((c) => `<option ${c === S.f.actClinic ? "selected" : ""}>${esc(c)}</option>`).join("")}</select></div>`;
    h += `<div class="btn-row" style="margin-bottom:10px">
      <button class="btn sm primary" data-act="mailArietis" ${vendor.length ? "" : "disabled"}>Email Arietis · ${vendor.length} at/past SLA</button>
      <button class="btn sm" data-act="remindOwners" ${owners.size ? "" : "disabled"}>Remind owners · ${owners.size}</button></div>`;
    const secOrSkip = (title, arr, extra, empty) => (S.f.find && !arr.length ? "" : sec(title, arr.length, extra) + list(arr, {}, empty));
    if (S.f.find && !ts.length) return h + list([], {});
    h += secOrSkip("Past SLA", crit, breached.length ? `<span class="muted">${breached.length} waiting on Arietis</span>` : "", "Nothing past SLA.");
    h += secOrSkip("Due today / update needed", warn, "", "Nothing due today.");
    h += secOrSkip("Housekeeping", info, "", "No housekeeping items.");
    return h;
  }

  // ------------------------------------------------------------------ ops escalations
  function opsSets() {
    const pool = S.tickets.filter((t) => inScope(t) && matches(t, S.f.find));
    const today = Math.floor(S.ctx.now);
    const queue = BE.opsQueue(pool);
    return {
      queue,
      due: queue.filter((t) => t.followUpDue !== null && Math.floor(t.followUpDue) <= today),
      unreviewed: queue.filter((t) => t.reviewedAt === null),
      suggest: BE.sortForAttention(pool.filter((t) => t.isOpen && !t.opsEscalated && BE.normKey(t.opsFlag) !== "cleared" && t.topSev === 1)),
      closed: pool.filter((t) => t.opsEscalated && !t.isOpen).sort((a, b) => (b.resolution.at || 0) - (a.resolution.at || 0)),
    };
  }
  function viewOps() {
    const o = opsSets();
    if (["queue", "due", "unreviewed", "suggest", "closed"].indexOf(S.f.ops) < 0) S.f.ops = "queue";
    const cur = o[S.f.ops];
    let h = `<div class="toolbar">${scopeSel()}${seg("ops", [["queue", "Escalated", o.queue.length], ["due", "F/U due", o.due.length], ["unreviewed", "Unreviewed", o.unreviewed.length], ["suggest", "Suggest", o.suggest.length], ["closed", "Closed", o.closed.length]], S.f.ops)}</div>`;
    if (S.f.ops !== "suggest" && S.f.ops !== "closed") {
      h += `<div class="btn-row" style="margin-bottom:10px"><button class="btn sm primary" data-act="reviewStart" ${cur.length ? "" : "disabled"}>Review with Arietis · ${cur.length}</button><button class="btn sm" data-act="opsExport" ${cur.length ? "" : "disabled"}>Export for Arietis</button></div>`;
    }
    if (S.f.ops === "suggest") h += `<p class="muted" style="margin:0 2px 8px">Open tickets past SLA that aren't escalated. Open one and tap <b>Escalate to ops</b>.</p>`;
    return h + list(cur, {}, S.f.ops === "queue" ? "No open ops escalations." : "Nothing here.");
  }
  function viewReview() {
    const R = S.review;
    const t = S.byId.get(R.ids[R.i]);
    if (!t) { S.review = null; return viewOps(); }
    const statusOpts = options(C.status).slice(); if (statusOpts.indexOf(t.status) < 0) statusOpts.unshift(t.status);
    const kv = (k, v, cls) => `<dt>${esc(k)}</dt><dd class="${cls || ""}">${v}</dd>`;
    const sla = (s) => `<span class="sla ${s.state === "breached" || s.state === "late" ? "bad" : s.state === "due-today" ? "warn" : s.state === "met" || s.state === "met-nodate" ? "ok" : "neutral"}">${esc(BE.slaText(s))}</span>`;
    return `<div class="review">
      <div class="review-bar"><b>Review with Arietis</b><span class="muted tnum">${R.i + 1} of ${R.ids.length}</span><span class="sp"></span>
        <button class="btn sm" data-act="reviewPrev" ${R.i ? "" : "disabled"} aria-label="Previous">←</button><button class="btn sm" data-act="reviewNext" ${R.i < R.ids.length - 1 ? "" : "disabled"} aria-label="Next">→</button><button class="btn sm" data-act="reviewExit">Exit</button></div>
      <div class="share-note">Share-safe view · internal work notes, staff emails and escalation reasons are hidden</div>
      <div class="card review-card">
        <div class="review-title"><span class="tid">#${t.id}</span><span class="mono">${esc(t.patient || "—")}</span><span class="mono muted">MRN ${esc(t.mrn || "—")}</span><span class="sp"></span><span class="chip ${stageClass(t)}">${esc(t.stage)}</span></div>
        <dl class="kv">${kv("Clinic", esc(t.clinic))}${kv("Issue", esc(t.taskTypes.join("; ") || "—"))}${kv("Amount", esc(BE.fmtMoney(t.amount) || "—"), "tnum")}
          ${kv("Submitted", esc(BE.fmtDate(t.submitted)) + ` <span class="muted">· ${t.ageBD} business days open</span>`, "tnum")}${kv("Status", esc(t.status))}
          ${kv("Receipt (" + S.ctx.receiptDays + " BD)", sla(t.receipt))}${kv("Resolution (" + S.ctx.resolutionDays + " BD)", sla(t.resolution))}
          ${kv("Last Arietis reply", esc(BE.fmtDateTime(t.lastReplyAt) || "—"), "tnum")}${kv("Details", esc(t.notes || "—"))}</dl>
      </div>
      <form class="card" id="reviewForm" autocomplete="off"><h4>Outcome of this review</h4><div class="form">
        <div class="field full"><label for="rvStatus">Status</label><select id="rvStatus">${statusOpts.map((o) => `<option ${o === t.status ? "selected" : ""}>${esc(o)}</option>`).join("")}</select></div>
        <div class="field full"><label for="rvCommit">Arietis commitment</label><textarea id="rvCommit" rows="2" placeholder="What Arietis agreed to do">${esc(t.commitment)}</textarea></div>
        <div class="field"><label for="rvDue">Follow-up due</label><input type="date" id="rvDue" value="${isoDay(t.followUpDue)}"></div>
        <div class="field"><span class="lab">Last reviewed</span><div class="tnum" style="padding-top:6px">${esc(t.reviewedAt ? BE.fmtDateTime(t.reviewedAt) : "Never")}</div></div>
        <div class="field full"><label for="rvNotes">Review notes</label><textarea id="rvNotes" rows="2" placeholder="Visible on screen while sharing">${esc(t.reviewNotes)}</textarea></div>
      </div></form>
      <div class="savebar"><button class="btn sm danger" data-act="reviewClear" title="Remove from the ops escalation view">De-escalate</button><span class="sp"></span><button class="btn sm primary" data-act="reviewSave">${R.i < R.ids.length - 1 ? "Save & next" : "Save & finish"}</button></div>
    </div>`;
  }
  async function reviewSave(clear) {
    const R = S.review; const id = R.ids[R.i]; const t = S.byId.get(id);
    const due = $("#rvDue").value;
    const edits = [
      { col: C.status, value: $("#rvStatus").value },
      { col: C.commitment, value: BE.norm($("#rvCommit").value) },
      { col: C.followUpDue, value: due ? Math.floor(BE.toSerial(due)) : "" },
      { col: C.reviewNotes, value: BE.norm($("#rvNotes").value) },
      { col: C.reviewedAt, value: nowSerial() },
    ];
    if (BE.isClosedStatus(edits[0].value) && t.resolutionDate === null) edits.push({ col: C.resolutionDate, value: Math.floor(S.ctx.now) });
    if (clear) edits.push({ col: C.opsFlag, value: "Cleared" });
    if (!S.me) { pickMe(); toast("Choose your name first so the review is attributed.", true); return; }
    try {
      busy(true, "Saving…");
      await S.src.saveTicket(id, edits, S.me);
      toast((clear ? "De-escalated #" : "Saved review for #") + id);
      if (R.i < R.ids.length - 1) R.i++; else { S.review = null; toast("Review finished · " + R.ids.length + " tickets"); }
      await refresh(false, true, true);
      $("#view").scrollTop = 0;
    } catch (e) { busy(false); toast("Save failed: " + (e.message || e), true); }
  }
  function opsExport() {
    const cur = opsSets()[S.f.ops] || [];
    const headers = BE.ARIETIS_EXPORT_HEADERS;
    const rows = BE.arietisExportRows(cur);
    const tsv = [headers].concat(rows).map((r) => r.map((c) => String(c).replace(/[\t\n]/g, " ")).join("\t")).join("\n");
    const html = BE.hTable(headers, rows.map((r) => r.map((c) => esc(c))));
    sheet("Export for Arietis", `<p class="muted" style="margin:0 0 8px">${cur.length} tickets. Vendor-safe fields only: no work notes, review notes, staff emails or escalation reasons. Nothing is sent until you send it.</p>
      <div class="btn-row" style="margin-bottom:10px"><button class="btn sm primary" data-act="opsEmail">Email to Arietis</button><button class="btn sm" data-act="opsCopy">Copy table</button></div>
      <div class="tbl-wrap" style="max-height:50vh;overflow:auto">${html.replace('<table', '<table class="grid"')}</div>`);
    S._export = { tsv, html, cur };
  }
  async function escalate(t, reason) {
    if (!S.me) { pickMe(); toast("Choose your name first.", true); return; }
    try {
      busy(true, "Escalating…");
      await S.src.saveTicket(t.id, [{ col: C.opsFlag, value: "Yes" }, { col: C.opsReason, value: reason }, { col: C.opsBy, value: S.me }, { col: C.opsAt, value: nowSerial() }], S.me);
      toast("Escalated #" + t.id + " to ops");
      await refresh(false, true, true);
    } catch (e) { busy(false); toast("Failed: " + (e.message || e), true); }
  }

  function viewClinic() {
    const c = S.f.clinic;
    const all = S.tickets.filter((t) => t.clinic === c && matches(t, S.f.find));
    const st = BE.computeStats(all).total;
    const by = { open: all.filter((t) => t.isOpen), await: all.filter((t) => t.stage === "Awaiting receipt"), with: all.filter((t) => t.stage === "With Arietis"), call: all.filter((t) => t.stage === "Pending call"), closed: all.filter((t) => !t.isOpen), all: all };
    const email = S.clinicEmails.get(c.toLowerCase()) || "";
    const mineFirst = profile.clinics.filter((x) => S.clinics.indexOf(x) >= 0).concat(S.clinics.filter((x) => profile.clinics.indexOf(x) < 0));
    let h = `<div class="toolbar"><select data-filter="clinic" aria-label="Clinic" style="flex:1">${mineFirst.map((x) => `<option ${x === c ? "selected" : ""}>${esc(x)}${profile.clinics.indexOf(x) >= 0 ? " ★" : ""}</option>`).join("")}</select>
      <button class="btn sm" data-act="mailClinic" ${email && by.open.length ? "" : "disabled"} title="${esc(email)}">Email clinic</button></div>`;
    h += `<div class="toolbar">${seg("stage", [["open", "Open", by.open.length], ["await", "Awaiting", by.await.length], ["with", "Arietis", by.with.length], ["call", "Pt call", by.call.length], ["closed", "Closed", by.closed.length]], S.f.stage)}</div>`;
    const ts = S.f.stage === "closed" ? by.closed.sort((a, b) => (b.resolution.at || 0) - (a.resolution.at || 0)) : attn(by[S.f.stage] || by.open);
    h += list(ts, { owner: true }, "No tickets in this view.");
    return h;
  }

  function viewStats() {
    const days = { all: null, "30": 30, "90": 90, mtd: "mtd" }[S.f.range];
    let ts = S.tickets.filter(inScope);
    const today = Math.floor(S.ctx.now);
    if (days === "mtd") { const d = BE.serialToDate(today); const first = BE.serialFromParts(d.getUTCFullYear(), d.getUTCMonth() + 1, 1, 0, 0, 0); ts = ts.filter((t) => (t.submitted || 0) >= first); }
    else if (days) ts = ts.filter((t) => (t.submitted || 0) >= today - days);
    const st = BE.computeStats(ts);
    const T = st.total;
    let h = `<div class="toolbar">${scopeSel()}${seg("range", [["all", "All time"], ["90", "90 days"], ["30", "30 days"], ["mtd", "Month to date"]], S.f.range)}<span class="muted">by submit date · ${ts.length} tickets</span></div>`;
    h += `<div class="kpis">${kpi("Open", T.open)}${kpi("Past SLA", T.breachedOpen, T.breachedOpen ? "bad" : "")}${kpi("Receipt ≤" + S.ctx.receiptDays + " BD", pctTxt(T.rcptOnTime, T.rcptOnTime + T.rcptMiss))}${kpi("Resolved ≤" + S.ctx.resolutionDays + " BD", pctTxt(T.resOnTime, T.resOnTime + T.resMiss))}${kpi("Avg BD to receipt", avg(T.rcptBdSum, T.rcptBdN))}${kpi("Avg BD to resolve", avg(T.resBdSum, T.resBdN))}${kpi("$ open", BE.fmtMoney(T.amountOpen) || "$0")}${kpi("Svc recovery", T.serviceRecovery, T.serviceRecovery ? "warn" : "")}</div>`;
    const cls = (v, bad) => (v === 0 ? "zero" : bad ? "bad" : "");
    const pc = (a, b) => { const p = BE.pct(a, b); return p === null ? `<td class="zero">—</td>` : `<td class="${p < 60 ? "bad" : p < 85 ? "warn" : "good"}">${p}%</td>`; };
    const mix = (s) => { const tot = s.total || 1; return `<div class="mix" title="Awaiting ${s.stage[0]} · With Arietis ${s.stage[1]} · Pt call ${s.stage[2]} · Closed ${s.stage[3]}">${["a", "w", "c", "d"].map((k, i) => `<i class="${k}" style="width:${(s.stage[i] / tot) * 100}%"></i>`).join("")}</div>`; };
    const tr = (s, foot) => `<tr ${foot ? "" : `data-clinic="${esc(s.key)}"`}><td>${esc(foot ? "Total" : shortClinic(s.key))}</td><td>${s.open}</td><td class="${cls(s.stage[0])}">${s.stage[0]}</td><td class="${cls(s.stage[1] + s.stage[2])}">${s.stage[1] + s.stage[2]}</td><td class="${cls(s.breachedOpen, true)}">${s.breachedOpen}</td>${pc(s.rcptOnTime, s.rcptOnTime + s.rcptMiss)}${pc(s.resOnTime, s.resOnTime + s.resMiss)}<td>${avg(s.resBdSum, s.resBdN)}</td><td>${s.total}</td><td>${mix(s)}</td></tr>`;
    h += sec("By clinic", st.byClinic.length, `<span class="muted">tap a row to open</span>`);
    h += `<div class="tbl-wrap"><table class="grid"><thead><tr><th>Clinic</th><th>Open</th><th title="Awaiting Arietis receipt">Await</th><th title="With Arietis, incl. pending patient call">W/ Ari</th><th>Past SLA</th><th title="Receipt within SLA">Rcpt</th><th title="Resolved within SLA">Res</th><th title="Avg business days to resolve">Avg BD</th><th>Total</th><th>Mix</th></tr></thead>
      <tbody>${st.byClinic.map((s) => tr(s)).join("")}</tbody><tfoot>${tr(T, true)}</tfoot></table></div>
      <div class="legend"><span><i style="background:var(--amber)"></i>Awaiting receipt</span><span><i style="background:var(--blue)"></i>With Arietis</span><span><i style="background:var(--violet)"></i>Pending pt call</span><span><i style="background:var(--green);opacity:.55"></i>Closed</span></div>`;
    const count = (fn) => { const m = new Map(); ts.forEach((t) => fn(t).forEach((k) => k && m.set(k, (m.get(k) || 0) + 1))); return Array.from(m).sort((a, b) => b[1] - a[1]); };
    const bars = (title, arr) => { if (!arr.length) return ""; const mx = arr[0][1]; return `<div>${sec(title, null)}<div class="bars">${arr.slice(0, 8).map(([k, n]) => `<div class="bar-row"><span class="lbl" title="${esc(k)}">${esc(k)}</span><span class="track"><span class="fill" style="display:block;width:${(n / mx) * 100}%"></span></span><span class="n">${n}</span></div>`).join("")}</div></div>`; };
    h += `<div class="two">${bars("Task type", count((t) => t.taskTypes))}${bars("Open by owner", count((t) => (t.isOpen ? [t.ownerName || t.ownerEmail] : [])))}${bars("Source of error (closed)", count((t) => (t.isOpen ? [] : [t.errorSource || "Not recorded"])))}${bars("Outcome (closed)", count((t) => (t.isOpen ? [] : BE.splitList(t.outcome).length ? BE.splitList(t.outcome) : ["Not recorded"])))}${bars("Source of inquiry", count((t) => [t.source || "Not recorded"]))}${bars("Flags", count((t) => t.flags))}</div>`;
    return h;
  }

  function viewLookup() {
    const q = S.f.q.trim();
    let h = `<input class="search" type="search" id="q" placeholder="MRN, patient code, or ticket #" value="${esc(q)}" autocomplete="off" aria-label="Search by MRN, patient code, or ticket number">`;
    if (q.length < 2) return h + `<p class="muted" style="margin:10px 2px">Type 2+ characters. Matches every ticket for that patient, open and closed. Press <span class="kbd">/</span> to jump here.</p>`;
    const k = q.toLowerCase().replace(/^#/, "");
    const digits = /^\d+$/.test(k);
    const hits = S.tickets.filter((t) => (digits ? t.mrn.indexOf(k) >= 0 || String(t.id) === k : t.patient.toLowerCase().indexOf(k) >= 0 || t.mrn.indexOf(k) >= 0));
    if (!hits.length) return h + `<div class="empty" style="margin-top:10px">No tickets match “${esc(q)}”.</div>`;
    const groups = new Map();
    hits.forEach((t) => { const g = (t.mrn || "—") + "|" + (t.patient || "—"); if (!groups.has(g)) groups.set(g, []); groups.get(g).push(t); });
    h += `<div style="margin-top:10px">`;
    groups.forEach((ts, g) => {
      const [mrn, pt] = g.split("|");
      const open = ts.filter((t) => t.isOpen).length;
      const amt = ts.reduce((s, t) => s + (t.amount || 0), 0);
      h += `<div class="pgroup"><div class="pgroup-h"><b>MRN ${esc(mrn)}</b><span>${esc(pt)}</span><span class="sp"></span><span class="muted">${ts.length} ticket${ts.length > 1 ? "s" : ""} · ${open} open${amt ? " · " + BE.fmtMoney(amt) : ""}</span></div>${list(ts.sort((a, b) => (b.submitted || 0) - (a.submitted || 0)), { owner: true })}</div>`;
    });
    return h + `</div>`;
  }

  function viewRecaps() {
    const clinics = S.clinics.filter((c) => c !== "(No clinic)");
    if (!S.f.recapClinic || clinics.indexOf(S.f.recapClinic) < 0) S.f.recapClinic = profile.clinics.find((c) => clinics.indexOf(c) >= 0) || S.f.clinic || clinics[0] || "";
    const r = BE.buildClinicRecap(S.tickets, S.f.recapClinic, S.f.recapMode, S.ctx, S.settings.workbookUrl);
    const to = S.clinicEmails.get(S.f.recapClinic.toLowerCase()) || "";
    S._recap = { r, to };
    return `<div class="toolbar"><select data-filter="recapClinic" aria-label="Clinic" style="flex:1">${clinics.map((c) => `<option ${c === S.f.recapClinic ? "selected" : ""}>${esc(c)}</option>`).join("")}</select>
      ${seg("recapMode", [["am", "AM · open"], ["pm", "PM · closed today"]], S.f.recapMode)}</div>
      <div class="btn-row" style="margin-bottom:8px"><button class="btn sm primary" data-act="recapMail" ${to && r.count ? "" : "disabled"}>Email clinic</button><button class="btn sm" data-act="recapCopy">Copy template</button></div>
      <p class="muted" style="margin:0 2px 8px">To: ${esc(to || "no clinic email in REF")} · ${esc(r.subject)} · ${r.count} ticket${r.count === 1 ? "" : "s"}${S.f.recapMode === "am" ? ` · ${r.needsFollowUp} need follow-up · ${r.resolutionOverdue} resolution overdue` : ""}</p>
      <iframe class="recapframe" title="Recap preview" srcdoc="${esc(r.html)}"></iframe>
      <p class="muted" style="margin:8px 2px">Preview only. Includes patient code and MRN, so send it only to the clinic's own inbox. The scheduled flow sends these automatically once it's on.</p>`;
  }
  function viewActivity() {
    const q = S.f.actq.trim().toLowerCase().replace(/^#/, "");
    const rows = S.activity.filter((a) => !q || String(a.id) === q || (a.field + " " + a.newVal + " " + a.by + " " + a.source).toLowerCase().indexOf(q) >= 0)
      .sort((a, b) => b.at - a.at).slice(0, 250);
    let h = `<input class="search find" type="search" id="actq" placeholder="Filter by ticket #, field, person or action" value="${esc(S.f.actq)}" autocomplete="off" aria-label="Filter activity">`;
    if (!rows.length) return h + `<div class="empty">No activity ${q ? "matches “" + esc(q) + "”" : "logged yet"}.</div>`;
    h += `<div class="list">${rows.map((a) => `<button class="arow" data-open="${a.id || ""}" ${a.id ? "" : "disabled"}><span class="t tnum">${esc(BE.fmtDateTime(a.at))}</span><span class="tid">${a.id ? "#" + a.id : ""}</span><span class="what"><b>${esc(a.field || a.source)}</b>${a.newVal || a.oldVal ? ` <span class="muted">${esc(a.oldVal || "—")} →</span> ${esc(a.newVal || "(blank)")}` : ""}</span><span class="who muted">${esc(a.by)}${a.source && a.source !== a.field ? " · " + esc(a.source) : ""}</span></button>`).join("")}</div>`;
    return h + `<p class="muted" style="margin:8px 2px">Latest ${rows.length} entries from the Activity Log${S.src.kind === "dataverse" ? "" : " sheet"}.</p>`;
  }

  function viewNew() {
    const url = (S.settings.formUrl || DEFAULT_FORM);
    const embed = url + (url.indexOf("embed=") < 0 ? (url.indexOf("?") < 0 ? "?" : "&") + "embed=true" : "");
    return `<div class="toolbar"><span class="muted" style="flex:1">${S.src.kind === "dataverse" ? "Submit below. The new ticket appears here within about a minute." : "Submit below. New tickets land in Raw Data, then sync to Master."}</span>
      <button class="btn sm primary" data-act="syncNow">Sync now</button><button class="btn sm" data-act="openForm">Open in browser</button></div>
      <iframe class="formframe" title="Patient Billing Escalation form" src="${esc(embed)}" allowfullscreen></iframe>`;
  }

  // ------------------------------------------------------------------ detail
  const EDIT = [
    { col: C.status, type: "select", full: true },
    { col: C.followUpAt, type: "date", label: "Last follow-up (you)" },
    { col: C.outreachDate, type: "date", label: "Patient outreach" },
    { col: C.resolutionDate, type: "date", label: "Date of resolution", req: true },
    { col: C.urgency, type: "select", label: "Urgency" },
    { col: C.errorSource, type: "select", label: "Source of error", req: true },
    { col: C.falseVerif, type: "select", label: "False verification by Arietis?", req: true },
    { col: C.serviceRecovery, type: "select", label: "Service recovery flag" },
    { col: C.ehr, type: "check", label: "EHR task created" },
    { col: C.outcome, type: "multi", label: "Outcome", full: true, req: true },
    { col: C.cc, type: "text", label: "CC (emails, ; separated)", full: true },
    { col: C.workNotes, type: "area", label: "Work notes (internal)", full: true },
  ];
  const DETAILS_EDIT = [
    { col: C.clinic, type: "select", label: "Clinic" }, { col: C.dept, type: "select", label: "Department" },
    { col: C.patient, type: "text", label: "Patient" }, { col: C.mrn, type: "text", label: "MRN" },
    { col: C.source, type: "select", label: "Source of inquiry" }, { col: C.amount, type: "text", label: "Amount" },
    { col: C.taskType, type: "multi", label: "Task type", full: true }, { col: C.notes, type: "area", label: "Notes", full: true },
    { col: C.ownerEmail, type: "text", label: "Assigned to (email)" }, { col: C.sentAt, type: "date", label: "Submitted to vendor" },
    { col: C.convId, type: "text", label: "Vendor conversation ID (verified)", full: true },
  ];
  function rawVal(t, col) {
    switch (col) {
      case C.status: return t.status; case C.firstReplyDate: return t.firstReplyDate; case C.outreachDate: return t.outreachDate;
      case C.resolutionDate: return t.resolutionDate; case C.urgency: return t.urgency; case C.errorSource: return t.errorSource;
      case C.falseVerif: return t.falseVerif; case C.serviceRecovery: return t.serviceRecovery; case C.ehr: return t.ehr;
      case C.outcome: return t.outcome; case C.cc: return t.cc.join("; "); case C.workNotes: return t.workNotes;
      case C.clinic: return t.clinic === "(No clinic)" ? "" : t.clinic; case C.dept: return t.dept; case C.patient: return t.patient;
      case C.mrn: return t.mrn; case C.source: return t.source; case C.amount: return t.amount === null ? "" : t.amount;
      case C.taskType: return t.taskTypes.join(";"); case C.notes: return t.notes;
      case C.followUpAt: return t.followUpAt; case C.ownerEmail: return t.assignee; case C.sentAt: return t.sentAt; case C.convId: return t.convId;
      case C.commitment: return t.commitment; case C.followUpDue: return t.followUpDue; case C.reviewNotes: return t.reviewNotes;
      default: return "";
    }
  }
  function fieldHtml(t, f) {
    const v = rawVal(t, f.col);
    const id = "f_" + f.col.replace(/\W+/g, "_");
    const label = esc(f.label || f.col.replace(/:$/, ""));
    let input = "";
    if (f.type === "select") {
      const opts = options(f.col).slice();
      if (v && opts.indexOf(v) < 0) opts.unshift(v);
      input = `<select id="${id}" data-col="${esc(f.col)}"><option value="">—</option>${opts.map((o) => `<option ${o === v ? "selected" : ""}>${esc(o)}</option>`).join("")}</select>`;
    } else if (f.type === "date") {
      input = `<input type="date" id="${id}" data-col="${esc(f.col)}" value="${isoDay(v)}">`;
    } else if (f.type === "check") {
      input = `<label class="checks"><input type="checkbox" id="${id}" data-col="${esc(f.col)}" ${v ? "checked" : ""}> Yes</label>`;
    } else if (f.type === "multi") {
      const cur = BE.splitList(v);
      const opts = options(f.col).slice();
      cur.forEach((c) => { if (opts.indexOf(c) < 0) opts.push(c); });
      input = `<div class="checks" id="${id}" data-col="${esc(f.col)}" data-multi="1">${opts.map((o) => `<label><input type="checkbox" value="${esc(o)}" ${cur.indexOf(o) >= 0 ? "checked" : ""}>${esc(o)}</label>`).join("")}</div>`;
    } else if (f.type === "area") {
      input = `<textarea id="${id}" data-col="${esc(f.col)}" rows="3">${esc(v)}</textarea>`;
    } else {
      input = `<input type="text" id="${id}" data-col="${esc(f.col)}" value="${esc(v)}">`;
    }
    return `<div class="field ${f.full ? "full" : ""}" data-field="${esc(f.col)}" data-req="${f.req ? 1 : 0}"><label for="${id}">${label}</label>${input}</div>`;
  }
  function readField(el) {
    const col = el.getAttribute("data-col");
    if (el.getAttribute("data-multi")) return { col, value: Array.from(el.querySelectorAll("input:checked")).map((i) => i.value).join(";") };
    if (el.type === "checkbox") return { col, value: el.checked };
    if (el.type === "date") return { col, value: el.value ? Math.floor(BE.toSerial(el.value)) : "" };
    if (col === C.mrn || col === C.amount) { const n = BE.toNum(el.value); return { col, value: n === null ? BE.norm(el.value) : n }; }
    return { col, value: BE.norm(el.value) };
  }
  function slaLine(label, s, days) {
    const txt = BE.slaText(s);
    const cls = s.state === "breached" || s.state === "late" ? "bad" : s.state === "due-today" ? "warn" : s.state === "met" || s.state === "met-nodate" ? "ok" : "neutral";
    return `<span class="lab">${label} <span class="muted">(${days} BD)</span></span><span class="tnum">due ${BE.weekdayName(s.due)} ${BE.fmtDate(s.due)}</span><span class="sla ${cls}">${esc(txt)}</span>`;
  }
  function viewDetail(t) {
    const hist = S.activity.filter((a) => a.id === t.id).sort((a, b) => b.at - a.at);
    const link = (u) => `<button class="linkish" data-url="${esc(u)}">${esc(decodeURIComponent(u.split("/").pop().split("?")[0]).slice(0, 48) || "Open")}</button>`;
    const attach = t.attachments ? t.attachments.split(/[;\s]+(?=https?:)/).filter(Boolean).map(link).join("<br>") : "—";
    let h = `<div class="detail-head"><button class="btn sm" data-act="back" aria-label="Back">←</button><span class="tid">#${t.id}</span><span class="chip ${stageClass(t)}">${esc(t.stage)}</span>${flagChips(t)}<span class="sp"></span><span class="muted tnum">${t.ageBD} BD old</span></div>`;
    h += `<div class="card"><div class="timeline">
      <span class="lab">Submitted</span><span class="tnum">${esc(BE.weekdayName(t.submitted || 0) + " " + BE.fmtDate(t.submitted) + " " + BE.fmtTime(t.submitted))}</span><span class="muted">${esc(t.ownerName)}</span>
      ${slaLine("Receipt", t.receipt, S.ctx.receiptDays)}${slaLine("Resolution", t.resolution, S.ctx.resolutionDays)}</div></div>`;
    if (t.actions.length) h += `<div class="card"><h4>Needs action</h4><ul class="actions-list">${t.actions.map((a) => `<li><span class="dot sev-${a.sev}"></span><span>${esc(a.text)} <span class="muted">· ${a.who === "arietis" ? "Arietis" : "owner"}</span></span></li>`).join("")}</ul></div>`;
    h += `<div class="btn-row" style="margin-bottom:10px"><button class="btn sm" data-act="mailTicketArietis">Email Arietis</button><button class="btn sm" data-act="mailTicketOwner">Email owner</button>${t.link ? `<button class="btn sm" data-url="${esc(t.link)}">Form response</button>` : ""}</div>`;
    if (t.opsEscalated) {
      h += `<div class="card ops-card"><h4>Ops escalation <span class="sp"></span><button type="button" class="btn sm danger" data-act="deescalate">De-escalate</button></h4>
        <div style="margin-bottom:8px">${t.opsWhy.map((w) => `<span class="flag amber">${esc(w)}</span>`).join(" ")}${t.opsBy ? ` <span class="muted">· ${esc(t.opsBy)} ${esc(BE.fmtShort(t.opsAt))}</span>` : ""}${t.reviewedAt ? ` <span class="muted">· reviewed ${esc(BE.fmtShort(t.reviewedAt))}</span>` : ""}</div>
        <div class="form">${[{ col: C.commitment, type: "text", label: "Arietis commitment", full: true }, { col: C.followUpDue, type: "date", label: "Follow-up due" }, { col: C.reviewNotes, type: "area", label: "Review notes", full: true }].map((f) => fieldHtml(t, f)).join("")}</div></div>`;
    } else if (t.isOpen) {
      h += `<div class="card"><h4>Ops escalation</h4><div class="toolbar" style="margin:0"><input type="text" id="escReason" placeholder="Reason (internal)" style="flex:1"><button type="button" class="btn sm" data-act="escalate">Escalate to ops</button></div></div>`;
    }
    h += `<form class="card" id="editForm" autocomplete="off"><h4>Update</h4><div class="form">${EDIT.map((f) => fieldHtml(t, f)).join("")}</div><div id="closeHint" class="hint" style="margin-top:6px"></div></form>`;
    h += `<div class="card"><h4>Escalation</h4><dl class="kv">
      <dt>Clinic</dt><dd>${esc(t.clinic)}</dd><dt>Patient</dt><dd class="mono">${esc(t.patient || "—")}</dd><dt>MRN</dt><dd class="mono">${esc(t.mrn || "—")}</dd>
      <dt>Department</dt><dd>${esc(t.dept || "—")}</dd><dt>Source</dt><dd>${esc(t.source || "—")}</dd><dt>Task type</dt><dd>${esc(t.taskTypes.join("; ") || "—")}</dd>
      <dt>Amount</dt><dd class="tnum">${esc(BE.fmtMoney(t.amount) || "—")}</dd><dt>Notes</dt><dd>${esc(t.notes || "—")}</dd><dt>Attachments</dt><dd>${attach}</dd>
      <dt>Owner</dt><dd>${esc(t.ownerEmail)}</dd></dl></div>`;
    h += `<details class="card collapse" id="detailsEdit"><summary><h4>Correct escalation details</h4></summary><p class="muted" style="margin:0 0 8px">Edits apply to Master only. Raw Data stays as the form submitted it.</p><div class="form">${DETAILS_EDIT.map((f) => fieldHtml(t, f)).join("")}</div></details>`;
    if (t.conflicts) h += `<div class="card ops-card"><h4>Form / Master conflict</h4><p style="margin:0">${esc(t.conflicts)}</p><p class="muted" style="margin:6px 0 0">The form and Master disagree, so Master was kept. Check Raw Data, set the right value below and save; that clears the flag.</p></div>`;
    h += `<div class="card"><h4>Arietis thread <span class="muted" style="text-transform:none;letter-spacing:0;font-weight:400">· filled by the inbox flow</span></h4><dl class="kv">
      <dt>Sent to Arietis</dt><dd class="tnum">${esc(BE.fmtDateTime(t.sentAt) || "—")}</dd><dt>First reply</dt><dd class="tnum">${esc(BE.fmtDateTime(t.firstReplyAt) || (t.firstReplyDate ? BE.fmtDate(t.firstReplyDate) + " (manual)" : "—"))}</dd>
      <dt>Last reply</dt><dd class="tnum">${esc(BE.fmtDateTime(t.lastReplyAt) || "—")}</dd><dt>Replies</dt><dd class="tnum">${t.replyCount || 0}</dd>
      <dt>Conversation</dt><dd class="mono muted" style="font-size:11px">${esc(t.convId || "not linked yet")}</dd></dl></div>`;
    h += `<div class="card"><h4>History <span class="muted">${hist.length}</span></h4>${hist.length ? `<ul class="history">${hist.slice(0, 40).map((a) => `<li><span class="t">${esc(BE.fmtDateTime(a.at))}</span><span><b>${esc(a.field.replace(/:$/, ""))}</b> → ${esc(a.newVal || "(blank)")} <span class="muted">· ${esc(a.by)}${a.source ? " · " + esc(a.source) : ""}</span></span></li>`).join("")}</ul>` : `<div class="muted">No changes logged yet.${t.updBy ? " Last updated by " + esc(t.updBy) + " " + esc(BE.fmtDateTime(t.updAt)) + "." : ""}</div>`}</div>`;
    h += `<div class="savebar"><span id="dirtyMsg" class="muted">No changes</span><span class="sp"></span><button class="btn sm" data-act="revert" disabled>Revert</button><button class="btn sm primary" data-act="save" disabled>Save to Master <span class="kbd" style="color:inherit;border-color:currentColor">Ctrl+↵</span></button></div>`;
    return h;
  }
  function detailEdits() {
    const t = S.byId.get(S.detail);
    const els = Array.from(document.querySelectorAll("#view [data-col]"));
    const edits = [];
    els.forEach((el) => {
      const e = readField(el);
      const orig = rawVal(t, e.col);
      const setEq = (a, b) => { const x = BE.splitList(a).map((v) => v.toLowerCase()).sort().join("|"); const y = BE.splitList(b).map((v) => v.toLowerCase()).sort().join("|"); return x === y; };
      const same = e.col === C.ehr ? !!orig === !!e.value : el.getAttribute("data-multi") ? setEq(orig || "", e.value) : BE.cellEq(orig === null ? "" : orig, e.value);
      el.closest(".field").classList.toggle("dirty", !same);
      if (!same) edits.push(e);
    });
    return edits;
  }
  function onDetailInput(ev) {
    const t = S.byId.get(S.detail);
    const el = ev && ev.target && ev.target.closest ? ev.target.closest("[data-col]") : null;
    if (el && el.getAttribute("data-col") === C.status && ev.type === "change") {
      const v = el.value;
      const today = isoDay(S.ctx.now);
      const set = (col, val) => { const x = document.querySelector(`#view [data-col="${CSS.escape(col)}"]`); if (x && !x.value) x.value = val; };
      if (BE.isClosedStatus(v)) set(C.resolutionDate, today);
    }
    const edits = detailEdits();
    const statusEl = document.querySelector(`#view [data-col="${CSS.escape(C.status)}"]`);
    const closing = statusEl && BE.isClosedStatus(statusEl.value);
    const missing = [];
    document.querySelectorAll('#view .field[data-req="1"]').forEach((f) => {
      const inp = f.querySelector("[data-col]");
      const val = readField(inp).value;
      const col = inp.getAttribute("data-col");
      const need = closing && BE.isBlank(val) && !(col === C.outcome && statusEl.value === BE.STATUS.closedNoAction);
      f.classList.toggle("req", need);
      if (need) missing.push(f.querySelector("label").textContent.replace(" · to close", ""));
    });
    $("#closeHint").textContent = missing.length ? "To close cleanly, also fill: " + missing.join(", ") + ". You can still save." : "";
    $("#dirtyMsg").textContent = edits.length ? edits.length + " unsaved change" + (edits.length > 1 ? "s" : "") : "No changes";
    document.querySelector('[data-act="save"]').disabled = !edits.length;
    document.querySelector('[data-act="revert"]').disabled = !edits.length;
  }
  async function saveDetail() {
    const edits = detailEdits();
    if (!edits.length) return;
    if (!S.me) { pickMe(); toast("Choose your name first so the change is attributed.", true); return; }
    const btn = document.querySelector('[data-act="save"]');
    btn.disabled = true;
    try {
      busy(true, "Saving…");
      const r = await S.src.saveTicket(S.detail, edits, S.me);
      toast(r.changed ? "Saved #" + S.detail + " · " + r.changed + " field" + (r.changed > 1 ? "s" : "") + " updated" : "No changes to save");
      await refresh(false, true, true);
    } catch (e) { busy(false); btn.disabled = false; toast("Save failed: " + (e.message || e), true); }
  }

  // ------------------------------------------------------------------ email
  function copyRich(html, text) {
    try {
      if (navigator.clipboard && window.ClipboardItem) {
        navigator.clipboard.write([new ClipboardItem({ "text/html": new Blob([html], { type: "text/html" }), "text/plain": new Blob([text], { type: "text/plain" }) })]).catch(() => copyFallback(html));
        return;
      }
    } catch (e) { /* fallthrough */ }
    copyFallback(html);
  }
  function copyFallback(html) {
    const d = document.createElement("div");
    d.contentEditable = "true"; d.style.cssText = "position:fixed;left:-9999px;top:0"; d.innerHTML = html;
    document.body.appendChild(d);
    const r = document.createRange(); r.selectNodeContents(d);
    const sel = window.getSelection(); sel.removeAllRanges(); sel.addRange(r);
    try { document.execCommand("copy"); } catch (e) { /* ignore */ }
    sel.removeAllRanges(); d.remove();
  }
  function htmlToText(html) {
    const d = document.createElement("div");
    d.innerHTML = html.replace(/<\/(p|h\d|tr|div)>/gi, "\n").replace(/<br\s*\/?>/gi, "\n").replace(/<\/t[dh]>/gi, "\t");
    return d.textContent.replace(/\t\n/g, "\n").replace(/\n{3,}/g, "\n\n").trim();
  }
  function compose(m) {
    const text = m.text || htmlToText(m.html || "");
    copyRich(m.html || esc(text).replace(/\n/g, "<br>"), text);
    let body = text;
    const enc = (s) => encodeURIComponent(s || "");
    if (enc(body).length > 1800) body = "[Formatted summary copied — press Ctrl+V here]\n";
    const url = "https://outlook.office.com/mail/deeplink/compose?to=" + enc(m.to) + (m.cc ? "&cc=" + enc(m.cc) : "") + "&subject=" + enc(m.subject) + "&body=" + enc(body);
    S.src.openUrl(url);
    toast("Draft opened in Outlook. Formatted copy is on your clipboard.");
  }
  function ticketBlock(t) {
    return `Ticket #${t.id} — submitted ${BE.fmtDate(t.submitted)} (${t.ageBD} business days ago)\nPatient: ${t.patient || "—"}   MRN: ${t.mrn || "—"}   Clinic: ${t.clinic}\nIssue: ${t.taskTypes.join("; ") || "—"}${t.amount ? " — " + BE.fmtMoney(t.amount) : ""}\nNotes: ${t.notes || "—"}\nCurrent status: ${t.status}`;
  }
  function mailTicketArietis(t) {
    const cc = [S.settings.billingInbox, S.clinicEmails.get(t.clinic.toLowerCase()), t.ownerEmail].concat(t.cc).filter(Boolean);
    const due = t.receipt.at === null && t.receipt.state !== "met-nodate"
      ? `Receipt confirmation was due ${BE.fmtDate(t.receipt.due)} (${S.ctx.receiptDays} business days).`
      : `Resolution is due ${BE.fmtDate(t.resolution.due)} (${S.ctx.resolutionDays} business days).`;
    compose({
      to: S.settings.arietisEmail, cc: Array.from(new Set(cc)).join(";"),
      subject: `[BE-${t.id}] Status request — patient ${t.patient || ""}${t.mrn ? " (MRN " + t.mrn + ")" : ""}`,
      text: `Hello Arietis team,\n\nFollowing up on the MHS patient billing escalation below. ${due} Please confirm receipt and send an update.\n\n${ticketBlock(t)}\n\nThank you,\n${meName() || "MHS Billing"}`,
    });
  }
  function mailTicketOwner(t) {
    compose({
      to: t.assignee || t.ownerEmail, cc: [t.ownerEmail !== t.assignee ? t.ownerEmail : ""].concat(t.cc).filter(Boolean).join(";"),
      subject: `Billing escalation #${t.id} — ${t.actions.length ? "action needed" : "status"}`,
      text: `Hi ${(t.ownerName || "").split(" ")[0] || "there"},\n\n${t.actions.length ? "Escalation #" + t.id + " needs attention:\n" + t.actions.map((a) => "• " + a.text).join("\n") : "Quick check-in on escalation #" + t.id + "."}\n\n${ticketBlock(t)}\nReceipt SLA: ${BE.slaText(t.receipt)}   Resolution SLA: ${BE.slaText(t.resolution)}\n\nPlease update it in the MHS Billing Tickets add-in.\n\nThanks,\n${meName()}`,
    });
  }
  function mailArietisBulk() {
    let ts = S.tickets.filter((t) => t.isOpen && t.actions.some((a) => a.who === "arietis") && matches(t, S.f.find));
    if (S.f.actWho === "mine" && S.me) ts = ts.filter((t) => t.ownerEmail === S.me || t.cc.indexOf(S.me) >= 0);
    if (S.f.actClinic) ts = ts.filter((t) => t.clinic === S.f.actClinic);
    const e = BE.buildArietisFollowup(ts, S.settings, S.ctx, meName());
    if (!e) { toast("Nothing at or past SLA."); return; }
    compose({ to: e.to, cc: e.cc, subject: e.subject, html: e.html });
  }
  function remindOwners() {
    const by = new Map();
    S.tickets.filter((t) => inScope(t) && t.actions.some((a) => a.who === "owner")).forEach((t) => { const k = t.assignee || t.ownerEmail; if (!by.has(k)) by.set(k, []); by.get(k).push(t); });
    const rows = Array.from(by).sort((a, b) => b[1].length - a[1].length).map(([email, ts]) => `<div class="trow" style="display:flex;align-items:center;gap:8px"><div style="flex:1;min-width:0"><b>${esc(S.people.get(email) || email)}</b><div class="muted" style="font-size:12px">${ts.length} ticket${ts.length > 1 ? "s" : ""} · ${ts.slice(0, 6).map((t) => "#" + t.id).join(", ")}${ts.length > 6 ? "…" : ""}</div></div><button class="btn sm" data-act="mailOwnerList" data-email="${esc(email)}">Compose</button></div>`).join("");
    sheet("Remind owners", `<p class="muted" style="margin:0 0 8px">One draft per owner listing their open action items. Daily emails go out automatically once the flows are on.</p><div class="list">${rows || '<div class="empty">No owner actions.</div>'}</div>`);
  }
  function mailOwnerList(email) {
    const ts = BE.sortForAttention(S.tickets.filter((t) => (t.assignee || t.ownerEmail) === email && t.actions.some((a) => a.who === "owner")));
    const name = (S.people.get(email) || "").split(" ")[0] || "there";
    const rows = ts.map((t) => [`<b>#${t.id}</b>`, esc(t.clinic), esc([t.patient, t.mrn].filter(Boolean).join(" · ")), t.actions.filter((a) => a.who === "owner").map((a) => esc(a.text)).join("<br>")]);
    const html = `<div style="font-family:Segoe UI,Arial,sans-serif;font-size:13px"><p>Hi ${esc(name)},</p><p>These billing escalations need an update from you:</p>${BE.hTable(["Ticket", "Clinic", "Patient", "Action"], rows)}<p>Update them in the MHS Billing Tickets add-in.</p><p>Thanks,<br>${esc(meName())}</p></div>`;
    compose({ to: email, subject: `Billing escalations — ${ts.length} need your update`, html });
  }
  function mailClinic() {
    const c = S.f.clinic;
    const ts = S.tickets.filter((t) => t.clinic === c);
    const e = BE.buildClinicDigests(ts, S.clinicEmails, S.ctx, S.settings)[0];
    if (!e) { toast("No open tickets for " + c); return; }
    compose({ to: e.to, cc: S.settings.billingInbox, subject: e.subject, html: e.html });
  }

  // ------------------------------------------------------------------ overlays
  function sheet(title, body) {
    closeSheet();
    const o = document.createElement("div");
    o.className = "overlay"; o.id = "overlay";
    o.innerHTML = `<div class="sheet" role="dialog" aria-modal="true" aria-label="${esc(title)}"><div style="display:flex;align-items:center;margin-bottom:8px"><h3 style="margin:0">${esc(title)}</h3><span class="sp"></span><button class="icon-btn" data-act="closeSheet" aria-label="Close">✕</button></div>${body}</div>`;
    o.addEventListener("click", (e) => { if (e.target === o) closeSheet(); });
    document.body.appendChild(o);
    const f = o.querySelector("input,select,button:not([data-act=closeSheet])"); if (f) f.focus();
  }
  function closeSheet() { const o = $("#overlay"); if (o) o.remove(); }
  function pickMe() {
    if (S.src && S.src.currentUser && S.src.currentUser()) { toast("Signed in as " + meName() + " (from your Microsoft sign-in)"); return; }
    const emails = Array.from(S.people.keys()).sort();
    const clinics = S.clinics.filter((c) => c !== "(No clinic)");
    sheet("Your profile", `<form id="profileForm" class="form" autocomplete="off">
      <div class="field"><label for="pfName">Name</label><input id="pfName" name="name" value="${esc(profile.name)}" required></div>
      <div class="field"><label for="pfEmail">Email (the one Microsoft Forms records)</label><input id="pfEmail" name="email" type="email" list="pfEmails" value="${esc(profile.email)}" required><datalist id="pfEmails">${emails.map((e) => `<option value="${esc(e)}">`).join("")}</datalist></div>
      <div class="field full"><label for="pfRole">Role</label><input id="pfRole" name="role" value="${esc(profile.role)}" placeholder="e.g. Practice manager"></div>
      <div class="field full"><span class="lab">Assigned clinics <span class="muted" style="font-weight:400">— default scope for Action, Ops and Stats; none = all</span></span><div class="checks">${clinics.map((c) => `<label><input type="checkbox" name="clinics" value="${esc(c)}" ${profile.clinics.indexOf(c) >= 0 ? "checked" : ""}>${esc(shortClinic(c))}</label>`).join("") || '<span class="muted">Clinics load from REF once the workbook is open.</span>'}</div></div>
      ${S.src.kind === "dataverse" ? "" : `<div class="field full"><label for="pfUrl">Cloud workbook URL (for email links; saved to Settings)</label><input id="pfUrl" name="workbookUrl" type="url" value="${esc(S.settings.workbookUrl)}" placeholder="https://…"></div>`}
      <div class="field full"><button class="btn primary" type="submit">Save profile</button><span class="muted">Stored on this computer only. Workbook permissions still decide who can see what.</span></div></form>`);
    $("#profileForm").addEventListener("submit", async (e) => {
      e.preventDefault();
      const fd = new FormData(e.target);
      const url = String(fd.get("workbookUrl") || "").trim();
      if (url && !/^https:\/\//i.test(url)) { toast("Use an https:// workbook URL.", true); return; }
      profile = { name: String(fd.get("name")).trim(), email: String(fd.get("email")).trim().toLowerCase(), role: String(fd.get("role") || "").trim(), clinics: fd.getAll("clinics").map(String) };
      try { localStorage.setItem(PROFILE_KEY, JSON.stringify(profile)); } catch (err) { /* private mode */ }
      S.me = BE.normKey(profile.email); S.f.scope = "assigned";
      if (url !== (S.settings.workbookUrl || "") && S.src.setWorkbookUrl) { try { await S.src.setWorkbookUrl(url); S.settings.workbookUrl = url; } catch (err) { toast("Couldn't save the workbook URL: " + err.message, true); } }
      closeSheet(); render(); toast("Profile saved");
    });
  }
  async function openSettings() {
    let st = {};
    try { st = await S.src.status(); } catch (e) { /* ignore */ }
    const ok = (b) => (b ? `<span class="sla ok">✓</span>` : `<span class="sla bad">missing</span>`);
    sheet("Settings", `<div class="card"><h4>You</h4><div class="toolbar"><span style="flex:1">${esc(meName() || "Not set")} <span class="muted">${esc(S.me)}${profile.role ? " · " + esc(profile.role) : ""}${profile.clinics.length ? " · " + profile.clinics.length + " assigned clinic" + (profile.clinics.length > 1 ? "s" : "") : ""}</span></span><button class="btn sm" data-act="pickMe">Edit profile</button></div></div>
      ${S.src.kind === "dataverse" ? `<div class="card"><h4>Data</h4><dl class="kv"><dt>Store</dt><dd>Dataverse (Power Pages)</dd><dt>Tickets</dt><dd>${S.tickets.length}</dd><dt>Loaded</dt><dd>${esc(S.lastSync ? BE.fmtDateTime(S.lastSync) : "—")}</dd></dl><div class="btn-row" style="margin-top:8px"><button class="btn sm primary" data-act="syncNow">Refresh</button></div></div>` : `<div class="card"><h4>Workbook</h4><dl class="kv"><dt>Raw Data (form)</dt><dd>${ok(st.raw)} read-only, never written</dd><dt>Master</dt><dd>${ok(st.master)} ${S.tickets.length} tickets · all edits land here</dd><dt>Sync</dt><dd>Every 30 s while this pane is open · last ${esc(S.lastSync ? BE.fmtDateTime(S.lastSync) : "—")}</dd><dt>Merge rule</dt><dd>Form blanks never erase Master; disagreements are flagged, not overwritten</dd></dl>
      <div class="btn-row" style="margin-top:8px"><button class="btn sm primary" data-act="syncNow">Sync now</button><button class="btn sm" data-act="init">Initialize / refresh</button></div></div>`}
      <div class="card"><h4>SLA rules</h4><dl class="kv"><dt>Receipt</dt><dd>${S.ctx ? S.ctx.receiptDays : 2} business days from submission</dd><dt>Resolution</dt><dd>${S.ctx ? S.ctx.resolutionDays : 5} business days from submission</dd><dt>Stale</dt><dd>${S.ctx ? S.ctx.staleDays : 3} business days without activity</dd><dt>Holidays</dt><dd>${S.ctx ? S.ctx.hol.size : 0} on the calendar</dd></dl><p class="muted" style="margin:8px 0 0">${S.src.kind === "dataverse" ? "Change these in the Escalation Settings and Holidays tables." : "Change these on the Settings sheet (column A/B values; holidays in column D)."}</p></div>
      <div class="card"><h4>Preview my automated emails</h4><div class="btn-row"><button class="btn sm" data-act="preview" data-mode="open">Start of day</button><button class="btn sm" data-act="preview" data-mode="close">End of day</button><button class="btn sm" data-act="preview" data-mode="weekly">Weekly recap</button></div></div>
      <p class="muted">MHS Billing Tickets v${VERSION} · keys: <span class="kbd">/</span> lookup · <span class="kbd">R</span> sync · <span class="kbd">Esc</span> back</p>`);
  }
  function preview(mode) {
    const all = BE.buildDigests(S.tickets, S.activity, S.ctx, S.settings, mode);
    const mine = all.filter((e) => e.to === S.me);
    const e = mine[0] || all[0];
    if (!e) { toast("Nothing would be sent for this run."); return; }
    sheet(e.subject, `<p class="muted" style="margin:0 0 8px">To: ${esc(e.to)}${mine.length ? "" : " (you have no tickets in this run; showing another owner's)"} · ${all.length} emails total this run</p><iframe title="Email preview" srcdoc="${esc(e.html)}"></iframe>`);
  }

  // ------------------------------------------------------------------ setup / fatal
  function renderSetup(st) {
    $("#view").innerHTML = `<div class="setup"><h3>One-time setup</h3><p>Prepares this workbook. Raw Data and the Forms connection aren't touched.</p>
      <ul>${st.master ? "" : "<li><b>Master</b> becomes a table (<span class='mono'>tblMaster</span>) with 20 tracking columns (Arietis reply times, CC, EHR task, last updated…). Existing rows and edits are kept as-is. The blank placeholder rows below the data are cleared.</li>"}
      ${st.activity ? "" : "<li><b>Activity</b> sheet: an audit log of every change.</li>"}
      ${st.settings && st.holidays ? "" : "<li><b>Settings</b> sheet: SLA days, email recipients, holiday calendar (prefilled with standard US holidays, so edit it to match yours).</li>"}
      <li>Then new form submissions sync from Raw Data into Master.</li></ul>
      <button class="btn primary" data-act="runSetup">Prepare workbook</button></div>`;
    $("#app").setAttribute("aria-busy", "false");
  }
  async function runSetup() {
    try {
      busy(true, "Setting up…");
      const log = await S.src.setup(S.me);
      S.setupNeeded = false;
      closeSheet();
      await refresh(true);
      S.src.watch(debounce(() => refresh(false, true), 1500));
      toast(log.length ? log[0] + (log.length > 1 ? " (+" + (log.length - 1) + " more)" : "") : "Already set up");
      if (!S.me) pickMe();
    } catch (e) { busy(false); toast("Setup failed: " + (e.message || e), true); }
  }
  function renderFatal(msg) { $("#view").innerHTML = `<div class="setup"><h3>Can't load</h3><p>${esc(msg)}</p><button class="btn" data-act="reload">Try again</button></div>`; busy(false); }
  function showBanner(msg) { const b = $("#banner"); b.textContent = msg; b.hidden = false; }

  // ------------------------------------------------------------------ render
  function render() {
    if (!S.ready) return;
    const me = $("#btnMe");
    me.textContent = meName() || "Set name";
    me.classList.toggle("unset", !S.me);
    const mineNeeds = S.me ? S.tickets.filter((t) => t.ownerEmail === S.me && t.actions.length).length : 0;
    $("#bMine").textContent = mineNeeds || "";
    $("#bAction").textContent = S.tickets.filter((t) => t.topSev === 1).length || "";
    const bo = $("#bOps"); if (bo) bo.textContent = BE.opsQueue(S.tickets).length || "";
    document.querySelectorAll(".tabs [data-tab]").forEach((b) => b.setAttribute("aria-selected", String(!S.detail && b.dataset.tab === S.tab)));
    const v = $("#view");
    const keepScroll = v.scrollTop;
    let html = "";
    if (S.review && S.tab === "ops") html = viewReview();
    else if (S.detail !== null && S.byId.has(S.detail)) html = viewDetail(S.byId.get(S.detail));
    else {
      S.detail = null;
      html = { mine: viewMine, action: viewAction, ops: viewOps, clinic: viewClinic, recaps: viewRecaps, activity: viewActivity, stats: viewStats, lookup: viewLookup, new: viewNew }[S.tab]();
      if (["mine", "action", "ops", "clinic"].indexOf(S.tab) >= 0) html = `<input class="search find" type="search" id="find" placeholder="Filter by MRN, patient code or #" value="${esc(S.f.find)}" autocomplete="off" aria-label="Filter tickets by MRN, patient code or ticket number">` + html;
    }
    if (S.tab === "new" && !S.detail && v.querySelector(".formframe")) return; // don't reload the embedded form
    v.innerHTML = html;
    if (S.detail !== null) onDetailInput();
    v.scrollTop = S._resetScroll ? 0 : keepScroll;
    S._resetScroll = false;
  }
  function go(tab) { S.tab = tab; S.detail = null; S.review = null; S._resetScroll = true; store.set("tab", tab); render(); if (tab === "lookup") { const q = $("#q"); if (q) q.focus(); } }
  function openTicket(id) { S.prevScroll = $("#view").scrollTop; S.detail = id; S._resetScroll = true; render(); $("#view").focus(); }
  function back() {
    if (S.detail !== null && detailEdits().length && !confirmLeave()) return;
    S.detail = null; render(); $("#view").scrollTop = S.prevScroll || 0;
  }
  function confirmLeave() {
    // window.confirm is blocked in Office task panes — use a two-tap pattern instead.
    if (S._leaveArmed && Date.now() - S._leaveArmed < 4000) { S._leaveArmed = 0; return true; }
    S._leaveArmed = Date.now(); toast("Unsaved changes. Press back again to discard."); return false;
  }

  // ------------------------------------------------------------------ events
  document.addEventListener("click", (e) => {
    const tab = e.target.closest(".tabs [data-tab]");
    if (tab) { if (S.detail !== null && detailEdits().length && !confirmLeave()) return; go(tab.dataset.tab); return; }
    const op = e.target.closest("[data-open]");
    if (op) { openTicket(+op.dataset.open); return; }
    const sg = e.target.closest("[data-seg]");
    if (sg) { S.f[sg.dataset.seg] = sg.dataset.v; S._resetScroll = true; render(); return; }
    const cl = e.target.closest("[data-clinic]");
    if (cl) { S.f.clinic = cl.dataset.clinic; store.set("clinic", S.f.clinic); S.f.stage = "open"; go("clinic"); return; }
    const u = e.target.closest("[data-url]");
    if (u) { S.src.openUrl(u.dataset.url); return; }
    const a = e.target.closest("[data-act]");
    if (!a) return;
    const t = S.byId.get(S.detail);
    switch (a.dataset.act) {
      case "back": back(); break;
      case "findAll": S.f.q = S.f.find; go("lookup"); break;
      case "reviewStart": { const cur = opsSets()[S.f.ops] || []; S.review = { ids: cur.map((x) => x.id), i: 0 }; S._resetScroll = true; render(); break; }
      case "reviewPrev": S.review.i = Math.max(0, S.review.i - 1); S._resetScroll = true; render(); break;
      case "reviewNext": S.review.i = Math.min(S.review.ids.length - 1, S.review.i + 1); S._resetScroll = true; render(); break;
      case "reviewExit": S.review = null; render(); break;
      case "reviewSave": reviewSave(false); break;
      case "reviewClear": reviewSave(true); break;
      case "opsExport": opsExport(); break;
      case "opsCopy": copyRich(S._export.html, S._export.tsv); toast("Copied " + S._export.cur.length + " rows. Paste into Excel, Teams or an email."); break;
      case "opsEmail": { const e = BE.buildArietisReview(S._export.cur, S.settings, S.ctx, meName()); closeSheet(); compose({ to: e.to, cc: e.cc, subject: e.subject, html: e.html }); break; }
      case "escalate": escalate(t, BE.norm(($("#escReason") || {}).value || "")); break;
      case "deescalate": if (S.me) { S.src.saveTicket(t.id, [{ col: C.opsFlag, value: "Cleared" }], S.me).then(() => { toast("De-escalated #" + t.id); refresh(false, true, true); }); } else pickMe(); break;
      case "save": saveDetail(); break;
      case "revert": render(); break;
      case "pickMe": pickMe(); break;
      case "recapCopy": copyRich(S._recap.r.html, htmlToText(S._recap.r.html)); toast("Recap copied. Paste it into an Outlook email."); break;
      case "recapMail": compose({ to: S._recap.to, cc: S.settings.billingInbox, subject: S._recap.r.subject, html: S._recap.r.html }); break;
      case "fullscreen": MODE === "dialog" ? Office.context.ui.messageParent(JSON.stringify({ k: "close" })) : openFull(); break;
      case "init": closeSheet(); (async () => { try { busy(true, "Initializing…"); const l = await S.src.setup(S.me); toast(l.length ? l.join(", ") : "Workbook already up to date"); await refresh(true); } catch (err) { busy(false); toast(err.message, true); } })(); break;      case "closeSheet": closeSheet(); break;
      case "syncNow": closeSheet(); refresh(true); break;
      case "runSetup": runSetup(); break;
      case "reload": location.reload(); break;
      case "openForm": S.src.openUrl(S.settings.formUrl || DEFAULT_FORM); break;
      case "mailTicketArietis": mailTicketArietis(t); break;
      case "mailTicketOwner": mailTicketOwner(t); break;
      case "mailArietis": mailArietisBulk(); break;
      case "remindOwners": remindOwners(); break;
      case "mailOwnerList": mailOwnerList(a.dataset.email); break;
      case "mailClinic": mailClinic(); break;
      case "preview": preview(a.dataset.mode); break;
    }
  });
  document.addEventListener("input", (e) => {
    if (e.target.id === "actq") { S.f.actq = e.target.value; const pos = e.target.selectionStart; render(); const f = $("#actq"); f.focus(); f.setSelectionRange(pos, pos); return; }
    if (e.target.id === "find") { S.f.find = e.target.value; const pos = e.target.selectionStart; render(); const f = $("#find"); f.focus(); f.setSelectionRange(pos, pos); return; }
    if (e.target.id === "q") { S.f.q = e.target.value; const pos = e.target.selectionStart; render(); const q = $("#q"); q.focus(); q.setSelectionRange(pos, pos); return; }
    if (S.detail !== null && e.target.closest("#view")) onDetailInput(e);
  });
  document.addEventListener("change", (e) => {
    const f = e.target.getAttribute && e.target.getAttribute("data-filter");
    if (f) { S.f[f] = e.target.value; if (f === "clinic") store.set("clinic", e.target.value); S._resetScroll = true; render(); return; }
    if (S.detail !== null && e.target.closest("#view")) onDetailInput(e);
  });
  document.addEventListener("keydown", (e) => {
    const typing = /INPUT|TEXTAREA|SELECT/.test(document.activeElement.tagName);
    if ((e.ctrlKey || e.metaKey) && e.key === "Enter" && S.detail !== null) { e.preventDefault(); saveDetail(); return; }
    if (e.key === "Escape" && S.review && !$("#overlay")) { S.review = null; render(); return; }
    if (e.key === "Escape") { if ($("#overlay")) closeSheet(); else if (S.detail !== null) back(); return; }
    if (typing) return;
    if (e.key === "/") { e.preventDefault(); go("lookup"); }
    else if (e.key === "r" || e.key === "R") refresh(true);
  });
  $("#btnSync").addEventListener("click", () => refresh(true));
  $("#btnMe").addEventListener("click", pickMe);
  $("#btnSettings").addEventListener("click", openSettings);

  // ------------------------------------------------------------------ start
  if (typeof Office !== "undefined" && Office.onReady) Office.onReady(() => boot());
  else boot();
})();
