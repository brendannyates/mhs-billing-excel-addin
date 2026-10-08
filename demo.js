/* Demo data source — used only when the page is opened outside Excel (preview / training).
   Synthetic tickets; no real patient data. Same interface as ExcelSource. */
(function () {
  "use strict";
  const BE = window.BE;
  const now = () => { const d = new Date(); return (d.getTime() - d.getTimezoneOffset() * 60000) / 86400000 + 25569; };
  const CLINICS = [["Castro - CA", "clinic@example.com"], ["San Rafael - CA", "clinic@example.com"], ["Burlingame - CA", "clinic@example.com"], ["Oakland - CA", "clinic@example.com"], ["Walnut Creek - CA", "clinic@example.com"], ["Seattle - WA", "clinic@example.com"]];
  const PEOPLE = [["demo.user@mymhs.com", "Demo User"], ["alex.rivera@mymhs.com", "Alex Rivera"], ["sam.lee@mymhs.com", "Sam Lee"], ["jordan.kim@mymhs.com", "Jordan Kim"]];
  const TASKS = ["Balance Dispute (Patient Request)", "Duplicate Balances", "Refund Request", "Superbill/EOB Request (Specify Below)", "Not Billing Under Insurance"];
  const SOURCES = ["Clinic Pre-Check", "Patient - Luma", "Patient - In-Person", "Patient - Phone"];
  const ST = [BE.STATUS.submitted, BE.STATUS.confirmed, BE.STATUS.pendingCall, BE.STATUS.resolved, BE.STATUS.closedNoAction];
  let seed = 7;
  const rnd = () => ((seed = (seed * 16807) % 2147483647) / 2147483647);
  const pick = (a) => a[Math.floor(rnd() * a.length)];
  const code = () => Array.from({ length: 6 }, () => String.fromCharCode(65 + Math.floor(rnd() * 26))).join("");

  function build() {
    const n0 = now();
    const rows = [];
    for (let i = 1; i <= 48; i++) {
      const sub = Math.floor(n0) - Math.floor(rnd() * 40) + 0.35 + rnd() * 0.3;
      const p = pick(PEOPLE);
      const age = n0 - sub;
      let st = age < 1.5 ? ST[0] : pick(age > 14 ? [ST[3], ST[3], ST[4], ST[1], ST[0]] : ST);
      const confirmed = st !== ST[0];
      const closed = st === ST[3] || st === ST[4];
      const r = BE.MASTER_COLS.map(() => "");
      const set = (c, v) => { r[BE.MASTER_COLS.indexOf(c)] = v; };
      set(BE.COL.id, i); set(BE.COL.start, sub - 0.01); set(BE.COL.submitted, sub);
      set(BE.COL.email, p[0]); set(BE.COL.name, p[1]); set(BE.COL.option, "Log an Escalation");
      set(BE.COL.dept, "Clinic"); set(BE.COL.clinic, pick(CLINICS)[0]); set(BE.COL.patient, code());
      set(BE.COL.mrn, 100000 + Math.floor(rnd() * 90000)); set(BE.COL.urgency, rnd() < 0.08 ? "Critical" : rnd() < 0.3 ? "High" : "Normal");
      set(BE.COL.source, pick(SOURCES)); set(BE.COL.taskType, pick(TASKS)); set(BE.COL.amount, Math.round(rnd() * 1800));
      set(BE.COL.notes, "Demo ticket — patient reports balance after insurance payment posted.");
      set(BE.COL.status, st);
      const first = confirmed || rnd() < 0.2 ? sub + 0.5 + rnd() * 3 : "";
      if (first !== "" && first < n0) { set(BE.COL.firstReplyAt, first); set(BE.COL.lastReplyAt, Math.min(n0, first + rnd() * 3)); set(BE.COL.replyCount, 1 + Math.floor(rnd() * 3)); if (confirmed) set(BE.COL.firstReplyDate, Math.floor(first)); }
      if (closed) {
        set(BE.COL.resolutionDate, Math.floor(Math.min(n0, sub + 2 + rnd() * 8)));
        set(BE.COL.errorSource, pick(["Arietis", "MHS", "Arietis + MHS"])); set(BE.COL.outcome, pick(["Refund", "Rebill/Recode Claims", "Patient Plan Education"]));
        if (rnd() < 0.8) set(BE.COL.falseVerif, rnd() < 0.2 ? "Yes" : "No");
      }
      set(BE.COL.serviceRecovery, rnd() < 0.1 ? "Yes" : "");
      set(BE.COL.ehr, rnd() < 0.3);
      if (!closed && rnd() < 0.18) { set(BE.COL.opsFlag, "Yes"); set(BE.COL.opsReason, pick(["Patient complaint to leadership", "Repeat issue", "High balance"])); set(BE.COL.opsBy, "ops.lead@mymhs.com"); set(BE.COL.opsAt, Math.min(n0, sub + 3)); if (rnd() < 0.6) { set(BE.COL.commitment, pick(["Reprocess claim and refund overpayment", "Call patient to explain EOB", "Resubmit to primary insurance"])); set(BE.COL.followUpDue, Math.floor(n0) + Math.floor(rnd() * 6) - 2); set(BE.COL.reviewedAt, n0 - 2); } }
      set(BE.COL.updBy, p[0]); set(BE.COL.updAt, Math.min(n0, sub + rnd() * 4)); set(BE.COL.rawSynced, sub);
      rows.push(r);
    }
    return rows;
  }
  const ref = {
    headers: ["ClinicName", "ClinicEmail", "Department Submitting Ticket:", "Patient Primary Clinic:", "MRN:", "Urgency Level:", "Source of Inquiry", "Task Type", "Amount:", "Notes/Encounter#/Other Relevant Information:", "Status", "Date of first reply from Arietis:", "Date of Patient Outreach (if applicable):", "Date of Resolution:", "Was false verification of balance/charge given by Arietis?", "Patient Feedback/Service Recovery Flag:", "Source of Error:", "Outcome:"],
    rows: [],
  };
  const cols = { 0: CLINICS.map((c) => c[0]), 1: CLINICS.map((c) => c[1]), 2: ["Clinic", "Operations", "Care Navigation"], 3: CLINICS.map((c) => c[0]), 5: ["Normal", "High", "Critical"], 6: SOURCES.concat(["Patient - Portal", "Patient - Email"]), 7: TASKS.concat(["Patient Requests Callback", "Other"]), 10: ST, 14: ["Yes", "No"], 15: ["Yes", "No"], 16: ["Arietis", "MHS", "Arietis + MHS"], 17: ["Refund", "Rebill/Recode Claims", "Patient Plan Education", "BIDF Incorrect - Needed Correction"] };
  for (let i = 0; i < 8; i++) ref.rows.push(ref.headers.map((_, c) => (cols[c] && cols[c][i]) || ""));

  let master = null;
  let activity = [];
  const DemoSource = {
    kind: "demo",
    async status() { return { missingCols: [], master: true, raw: true, ref: true, activity: true, settings: true, holidays: true }; },
    async setup() { return []; },
    async addColumns(c) { return c; },
    async load() {
      if (!master) master = { headers: BE.MASTER_COLS.slice(), rows: build() };
      return { master, ref, holidays: [], settings: [], activity };
    },
    async sync() { return { appended: [], updated: 0, baselined: 0, duplicates: [], at: now() }; },
    async saveTicket(id, edits, by) {
      const r = master.rows.find((x) => x[0] === id);
      let n = 0;
      for (const e of edits) {
        const c = BE.MASTER_COLS.indexOf(e.col);
        if (c >= 0 && !BE.cellEq(r[c], e.value)) { activity.push([now(), id, e.col, BE.norm(r[c]), BE.norm(e.value), by, "Add-in"]); r[c] = e.value; n++; }
      }
      if (n) { r[BE.MASTER_COLS.indexOf(BE.COL.updBy)] = by; r[BE.MASTER_COLS.indexOf(BE.COL.updAt)] = now(); }
      return { changed: n };
    },
    async watch() {},
    openUrl(url) { window.open(url, "_blank", "noopener"); },
  };
  window.DemoSource = DemoSource;
})();
