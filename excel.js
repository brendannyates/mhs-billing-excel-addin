/* excel.js — every workbook read/write the add-in makes.
   Raw Data (the Forms-linked OfficeForms.Table) is READ ONLY. Edits go to Master only.
   Sync semantics are the repo's tested merge (core.js): blanks never erase Master, a changed form value
   only replaces Master when Master still equals the last observed form value, otherwise it is flagged
   as a conflict; add-in edits mark a field Master-owned in _SyncState. */
import { merge, filled, closed } from './core.js?v=5';

const BE = globalThis.BE;
const C = BE.COL;
const LOG = 'Activity Log';
const LOG_HEADERS = ['Timestamp', 'User', 'Action', 'Ticket Id', 'Field', 'Old value', 'New value', 'Details'];
// MHS-only link to the live workbook (requires an MHS Microsoft 365 sign-in). Initialize writes it to
// Settings "Workbook URL" when that cell is blank, so email links work without anyone typing it.
const DEFAULT_WORKBOOK_URL = 'https://mindfulhealthsolutions-my.sharepoint.com/:x:/g/personal/byates_mymhs_com/IQDxqsZTKhXFT7p_oSUG49IKAS_reuY_Z6SKWgNHnmnk3Eg?e=MURJfQ&nav=MTVfe0I1M0ZDMzgyLThCQkMtNDhEMS04NjkxLThDODNBOEM3RkQ2Rn0';
const SETTING_DEFAULTS = [
  ['Receipt SLA business days', 2], ['Resolution SLA business days', 5], ['Stale after business days', 3],
  ['Arietis email', 'patientbilling@arietishealth.com'], ['Billing inbox', 'billing@mindfulhealthsolutions.com'],
  ['Leadership recipients', ''], ['Clinic daily digest', 'No'], ['Arietis daily followup', 'No'], ['Digest include patient IDs', 'No'],
  ['Form URL', 'https://forms.cloud.microsoft/Pages/ResponsePage.aspx?id=8Sq2y8CkOEWdeGKInLNdYw1_rrxVxaFImLkzEWnt0GFUQUtJVFg2VzU5UEowM01JMUgzUUpKMTZPVi4u'],
];
const nowSerial = () => { const d = new Date(); return (d.getTime() - d.getTimezoneOffset() * 60000) / 86400000 + 25569; };
const show = (col, v) => (BE.DATE_COLS.includes(col) ? (BE.toSerial(v) === null ? '' : BE.fmtDateTime(BE.toSerial(v))) : BE.norm(v));
let snapshot = new Map(); // Id -> Master row object as last loaded (for "someone else changed this" checks)

async function used(ctx, name) {
  const ws = ctx.workbook.worksheets.getItemOrNullObject(name); await ctx.sync();
  if (ws.isNullObject) return null;
  const r = ws.getUsedRangeOrNullObject(true); r.load('values,rowCount,columnCount'); await ctx.sync();
  return { ws, values: r.isNullObject ? [] : r.values };
}
function appendLog(ctx, logValues, entries) {
  if (!entries.length || !logValues) return;
  const ws = ctx.workbook.worksheets.getItem(LOG);
  const start = Math.max(1, logValues.length);
  const rng = ws.getRangeByIndexes(start, 0, entries.length, LOG_HEADERS.length);
  rng.values = entries;
  ws.getRangeByIndexes(start, 0, entries.length, 1).numberFormat = entries.map(() => ['m/d/yyyy h:mm AM/PM']);
}
const logRow = (user, action, id, field, oldV, newV, details) => [nowSerial(), user || '', action, id === '' ? '' : Number(id), field || '', oldV || '', newV || '', details || ''];
// Excel turns 'YYYY-MM-DD' text into date serials; compare date columns by value so syncs don't rewrite them forever.
const DERIVED_DATES = [C.confirmDue, C.resolveDue];
function sameCell(col, a, b) {
  if (String(a ?? '') === String(b ?? '')) return true;
  if (BE.DATE_COLS.includes(col) || DERIVED_DATES.includes(col)) {
    const x = BE.toSerial(a), y = BE.toSerial(b);
    if (x !== null && y !== null) return Math.abs(x - y) < 1 / 86400;
  }
  return BE.cellEq(a ?? '', b ?? '');
}
function fmtFor(col) { return BE.DATE_ONLY_COLS.includes(col) ? 'm/d/yyyy' : BE.DATE_COLS.includes(col) ? 'm/d/yyyy h:mm AM/PM' : null; }

export const ExcelSource = {
  kind: 'excel',

  async status() { return { missingCols: [], master: true, raw: true, ref: true, activity: true, settings: true, holidays: true }; },
  async addColumns(c) { return c; },

  /** Initialize / refresh (idempotent): Master operational columns + table, _SyncState, Settings, Activity Log. */
  async setup(me) {
    return Excel.run(async (ctx) => {
      const log = [];
      const master = ctx.workbook.worksheets.getItem('Master');
      const range = master.getUsedRange(); range.load('values'); await ctx.sync();
      const h = range.values[0].map(String);
      for (const name of BE.EXTRA_COLS) if (!h.includes(name)) { master.getCell(0, h.length).values = [[name]]; h.push(name); log.push('Added column ' + name); }
      const tabs = ctx.workbook.worksheets; tabs.load('items/name'); await ctx.sync();
      const has = (n) => tabs.items.some((x) => x.name === n);
      if (!has('_SyncState')) { const s = tabs.add('_SyncState'); s.getRange('A1:D1').values = [['Id', 'Baseline JSON', 'Override Fields JSON', 'Conflicts JSON']]; s.visibility = 'Hidden'; log.push('Added _SyncState'); }
      if (!has('Settings')) {
        const s = tabs.add('Settings');
        s.getRange('A1:B6').values = [['Setting', 'Value'], ['Time zone', 'Pacific Standard Time'], ['Business open', '09:00'], ['Business close', '17:00'], ['Weekly recap day', 'Friday'], ['Workbook URL', DEFAULT_WORKBOOK_URL]];
        s.getRange('D1:D2').values = [['Holiday dates (YYYY-MM-DD)'], ['']];
        log.push('Added Settings');
      }
      if (!has(LOG)) { const s = tabs.add(LOG); s.getRange('A1:H1').values = [LOG_HEADERS]; s.freezePanes.freezeRows(1); log.push('Added Activity Log'); }
      await ctx.sync();
      // Settings: make sure every rules key exists (append missing rows under A:B; row 6 stays Workbook URL)
      const st = ctx.workbook.worksheets.getItem('Settings').getUsedRange(); st.load('values'); await ctx.sync();
      const keys = new Set(st.values.map((r) => String(r[0]).toLowerCase().replace(/[^a-z0-9]/g, '')));
      let next = Math.max(st.values.length, 6);
      const add = SETTING_DEFAULTS.filter(([k]) => !keys.has(k.toLowerCase().replace(/[^a-z0-9]/g, '')));
      if (add.length) { ctx.workbook.worksheets.getItem('Settings').getRangeByIndexes(next, 0, add.length, 2).values = add; next += add.length; }
      // Workbook URL: fill the built-in MHS link if the cell is blank (never overwrite a value someone set).
      const urlRow = st.values.findIndex((r) => String(r[0]).toLowerCase().replace(/[^a-z0-9]/g, '') === 'workbookurl');
      if (urlRow >= 0 && !String(st.values[urlRow][1] || '').trim()) {
        ctx.workbook.worksheets.getItem('Settings').getCell(urlRow, 1).values = [[DEFAULT_WORKBOOK_URL]];
        log.push('Set Workbook URL');
      }
      master.tables.load('items/name'); await ctx.sync();
      const tableRange = master.getRangeByIndexes(0, 0, range.values.length, h.length);
      if (!master.tables.items.length) { const t = master.tables.add(tableRange, true); t.name = 'BillingMaster'; }
      else master.tables.items[0].resize(tableRange);
      await ctx.sync();
      return log;
    });
  },

  async load() {
    return Excel.run(async (ctx) => {
      const m = await used(ctx, 'Master'), ref = await used(ctx, 'REF'), st = await used(ctx, 'Settings'), lg = await used(ctx, LOG);
      const headers = m.values[0].map(String);
      const rows = m.values.slice(1).filter((r) => filled(r[0]));
      snapshot = new Map(rows.map((r) => [String(r[0]), Object.fromEntries(headers.map((h, i) => [h, r[i] ?? '']))]));
      const sv = st ? st.values : [];
      return {
        master: { headers, rows },
        ref: ref && ref.values.length ? { headers: ref.values[0].map(String), rows: ref.values.slice(1) } : null,
        settings: sv.slice(1).map((r) => [r[0], r[1]]),
        holidays: sv.slice(1).map((r) => String(r[3] || '')).filter((x) => /^\d{4}-\d{2}-\d{2}$/.test(x)),
        activity: (lg ? lg.values.slice(1) : []).filter((r) => filled(r[0])).map((r) => [r[0], r[3], r[4] || r[2], r[5], r[6], r[1], r[2]]),
      };
    });
  },

  /** Raw Data -> Master (repo merge policy) + derived SLA columns from the shared rules + Activity Log. */
  async sync(user) {
    return Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getItem('Master'), state = ctx.workbook.worksheets.getItem('_SyncState');
      const mr = ws.getUsedRange(), rr = ctx.workbook.worksheets.getItem('Raw Data').getUsedRange(), sr = state.getUsedRange();
      mr.load('values'); rr.load('values'); sr.load('values'); await ctx.sync();
      const lg = await used(ctx, LOG), st = await used(ctx, 'Settings');
      const mh = mr.values[0].map(String), rh = rr.values[0].map(String), index = new Map();
      mr.values.slice(1).forEach((r, i) => { if (filled(r[0])) { if (index.has(String(r[0]))) throw Error('Duplicate Id ' + r[0] + ' on Master; sync stopped. Remove the extra row.'); index.set(String(r[0]), i + 1); } });
      const seen = new Set();
      for (const r of rr.values.slice(1)) { if (!filled(r[0])) continue; const id = String(r[0]); if (seen.has(id)) throw Error('Duplicate Id ' + id + ' in Raw Data; sync stopped.'); seen.add(id); }
      const states = new Map(sr.values.slice(1).filter((r) => filled(r[0])).map((r) => [String(r[0]), r]));
      const out = mr.values.map((r) => r.slice()); // full Master grid we will diff against
      const entries = [], appended = [];
      let merged = 0;
      for (const rawRow of rr.values.slice(1)) {
        if (!filled(rawRow[0])) continue;
        const id = String(rawRow[0]), r = Object.fromEntries(rh.map((h, i) => [h, rawRow[i] ?? '']));
        let pos = index.get(id);
        const m = pos === undefined ? {} : Object.fromEntries(mh.map((h, i) => [h, mr.values[pos][i] ?? '']));
        const old = states.get(id), baseline = old ? JSON.parse(old[1] || '{}') : {}, overrides = old ? JSON.parse(old[2] || '[]') : [];
        const res = merge(m, r, baseline); for (const key of overrides) res.next[key] = m[key] ?? '';
        // Master-owned fields (Portal rule): keep Master once filled; never a conflict. Never-merge fields keep Master.
        for (const key of BE.MASTER_OWNED_COLS) if (pos !== undefined && filled(m[key])) res.next[key] = m[key];
        for (const key of BE.NEVER_MERGE_COLS) if (pos !== undefined && key in m) res.next[key] = m[key];
        res.conflicts = res.conflicts.filter((k) => !BE.MASTER_OWNED_COLS.includes(k) && !BE.NEVER_MERGE_COLS.includes(k));
        if (closed(m)) res.next[C.status] = m[C.status];
        if (filled(m[C.firstReplyDate])) res.next[C.firstReplyDate] = m[C.firstReplyDate];
        const pending = new Set((old ? JSON.parse(old[3] || '[]') : []).filter((k) => !BE.MASTER_OWNED_COLS.includes(k) && !BE.NEVER_MERGE_COLS.includes(k))); res.conflicts.forEach((k) => pending.add(k));
        if (!filled(res.next[C.ownerEmail])) res.next[C.ownerEmail] = res.next[C.email] || '';
        if (closed(res.next) && !closed(m) && !filled(res.next[C.closedAt])) { res.next[C.status] = m[C.status] || BE.STATUS.submitted; pending.add('Status: (close in add-in after inbox reconciliation)'); }
        if (!filled(res.next[C.status])) res.next[C.status] = BE.STATUS.submitted;
        res.next[C.conflicts] = [...pending].join('; ');
        if (pos === undefined) {
          pos = out.length; out.push(mh.map(() => '')); index.set(id, pos); appended.push(Number(id));
          entries.push(logRow('Form', 'Imported', id, '', '', '', 'New form submission copied to Master'));
        } else {
          for (const key of BE.MERGEABLE_COLS) {
            if (!mh.includes(key) || String(res.next[key] ?? '') === String(m[key] ?? '')) continue;
            entries.push(logRow(r[C.email] || 'Form', 'Form edit', id, key.replace(/:$/, ''), show(key, m[key]), show(key, res.next[key]), 'Merged from Raw Data'));
            merged++;
          }
          res.conflicts.filter((k) => !(old && JSON.parse(old[3] || '[]').includes(k))).forEach((k) => entries.push(logRow('Form', 'Conflict', id, k.replace(/:$/, ''), show(k, m[k]), show(k, r[k]), 'Form and Master disagree; Master kept')));
        }
        mh.forEach((h, j) => { if (h in res.next) out[pos][j] = res.next[h] ?? ''; });
        states.set(id, [id, JSON.stringify(r), JSON.stringify(overrides), JSON.stringify([...pending])]);
      }
      // Derived columns for every Master row, using the shared rules (SLA from submission; action list)
      const sv = st ? st.values : [];
      const cfg = BE.readSettings(sv.slice(1).map((x) => [x[0], x[1]]));
      const ctxRules = { now: nowSerial(), hol: BE.holidaySet(sv.slice(1).map((x) => x[3]).filter((x) => /^\d{4}-\d{2}-\d{2}$/.test(String(x || '')))), receiptDays: cfg.receiptDays, resolutionDays: cfg.resolutionDays, staleDays: cfg.staleDays };
      const evaluated = BE.loadTickets(mh, out.slice(1), ctxRules);
      for (const t of evaluated) {
        const pos = index.get(String(t.id)); if (pos === undefined) continue;
        for (const [col, value] of BE.derivedColumns(t)) { const j = mh.indexOf(col); if (j >= 0) out[pos][j] = value; }
      }
      // Write only cells that changed; whole new rows in one call each
      const original = mr.values.length;
      for (let i = 1; i < out.length; i++) {
        if (i >= original) { ws.getRangeByIndexes(i, 0, 1, mh.length).values = [out[i]]; continue; }
        for (let j = 0; j < mh.length; j++) if (!sameCell(mh[j], out[i][j], mr.values[i][j])) ws.getCell(i, j).values = [[out[i][j] ?? '']];
      }
      if (states.size) state.getRangeByIndexes(1, 0, states.size, 4).values = [...states.values()];
      appendLog(ctx, lg && lg.values, entries);
      ws.tables.load('items/name'); await ctx.sync();
      if (ws.tables.items.length) ws.tables.items[0].resize(ws.getRangeByIndexes(0, 0, out.length, mh.length));
      await ctx.sync();
      return { appended, updated: merged, baselined: 0, duplicates: [], at: nowSerial() };
    });
  },

  /** Save edits from the add-in. Refuses if someone else changed an edited field since it was loaded. */
  async saveTicket(id, edits, by) {
    return Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getItem('Master'), state = ctx.workbook.worksheets.getItem('_SyncState');
      const range = ws.getUsedRange(), sr = state.getUsedRange(); range.load('values'); sr.load('values'); await ctx.sync();
      const lg = await used(ctx, LOG);
      const h = range.values[0].map(String), pos = range.values.findIndex((r, i) => i > 0 && String(r[0]) === String(id));
      if (pos < 0) throw Error('Ticket #' + id + ' is no longer on Master. Refresh.');
      const before = snapshot.get(String(id)) || {};
      const cur = Object.fromEntries(h.map((k, i) => [k, range.values[pos][i] ?? '']));
      const changes = edits.filter((e) => h.includes(e.col) && !(e.col === C.ehr ? BE.toBool(cur[e.col]) === BE.toBool(e.value) : BE.cellEq(cur[e.col], e.value)));
      for (const e of changes) if (!BE.cellEq(cur[e.col], before[e.col] ?? cur[e.col])) throw Error(e.col.replace(/:$/, '') + ' was changed by someone else. Refresh before saving.');
      if (!changes.length) return { changed: 0 };
      const next = { ...cur }; changes.forEach((e) => { next[e.col] = e.value; });
      const closing = closed(next) && !closed(cur);
      if (closed(cur) && String(next[C.status]) !== String(cur[C.status])) throw Error('Reopening a closed ticket needs a reviewed workflow. Status not changed.');
      if (closing && !filled(next[C.resolutionDate])) throw Error('Enter the date of resolution before closing.');
      const now = nowSerial(), sys = [[C.updBy, by], [C.updAt, now]];
      if (closing && !filled(cur[C.closedAt])) sys.push([C.closedAt, now]);
      for (const e of changes) {
        const j = h.indexOf(e.col), c = ws.getCell(pos, j);
        c.values = [[e.value]];
        const f = fmtFor(e.col); if (f && typeof e.value === 'number') c.numberFormat = [[f]];
      }
      for (const [col, v] of sys) { const j = h.indexOf(col); if (j >= 0) { const c = ws.getCell(pos, j); c.values = [[v]]; if (typeof v === 'number') c.numberFormat = [['m/d/yyyy h:mm AM/PM']]; } }
      // Master-owned fields: future form syncs never overwrite them
      const sp = sr.values.findIndex((r, i) => i > 0 && String(r[0]) === String(id));
      const old = sp < 0 ? [] : JSON.parse(sr.values[sp][2] || '[]');
      const override = [...new Set([...old, ...changes.map((x) => x.col)])];
      const stateRow = sp < 0 ? [String(id), '{}', JSON.stringify(override), '[]'] : [...sr.values[sp]];
      stateRow[2] = JSON.stringify(override); stateRow[3] = '[]';
      state.getRangeByIndexes(sp < 0 ? sr.values.length : sp, 0, 1, 4).values = [stateRow];
      const cj = h.indexOf(C.conflicts); if (cj >= 0) ws.getCell(pos, cj).values = [['']];
      appendLog(ctx, lg && lg.values, changes.map((e) => logRow(by, 'Edit', id, e.col.replace(/:$/, ''), show(e.col, cur[e.col]), show(e.col, e.value), 'Add-in')));
      await ctx.sync();
      return { changed: changes.length };
    });
  },

  async watch(cb) {
    try { await Excel.run(async (ctx) => { ctx.workbook.worksheets.getItem('Master').onChanged.add(async () => cb()); await ctx.sync(); }); } catch (e) { /* events unsupported */ }
  },
  openUrl(url) {
    try { if (Office.context.ui && Office.context.ui.openBrowserWindow) { Office.context.ui.openBrowserWindow(url); return; } } catch (e) { /* fall through */ }
    window.open(url, '_blank', 'noopener');
  },
};
globalThis.ExcelSource = ExcelSource;
globalThis.beNowSerial = nowSerial;
ExcelSource.setWorkbookUrl = async (url) => Excel.run(async (ctx) => {
  ctx.workbook.worksheets.getItem('Settings').getRange('A6:B6').values = [['Workbook URL', url]];
  await ctx.sync();
});
