// Runs the add-in's real excel.js sync against an in-memory workbook (synthetic data only).
import test from 'node:test'; import assert from 'node:assert/strict'; import fs from 'node:fs'; import vm from 'node:vm';
vm.runInThisContext(fs.readFileSync(new URL('../rules.js', import.meta.url), 'utf8'));
const { ExcelSource } = await import('../excel.js');
const BE = globalThis.BE;
function makeBook(sheets) {
  const grid = (n) => sheets[n];
  const rng = (n, r, c, nr, nc) => ({
    set values(v) { v.forEach((row, i) => { grid(n)[r + i] ??= []; row.forEach((x, j) => { grid(n)[r + i][c + j] = x; }); }); },
    set numberFormat(_) {}, load() {},
  });
  const sheet = (n) => ({
    isNullObject: false,
    getUsedRange() { return this.getUsedRangeOrNullObject(); },
    getUsedRangeOrNullObject() { const g = grid(n); let last = -1; g.forEach((r, i) => { if ((r || []).some((x) => x !== '' && x != null)) last = i; }); const w = Math.max(...g.map((r) => (r || []).length), 1); return { isNullObject: false, load() {}, values: g.slice(0, last + 1).map((r) => Array.from({ length: w }, (_, j) => (r || [])[j] ?? '')) }; },
    getRangeByIndexes: (r, c, nr, nc) => rng(n, r, c, nr, nc), getCell: (r, c) => rng(n, r, c, 1, 1), getRange: () => rng(n, 0, 0, 1, 1),
    tables: { items: [], load() {} },
  });
  return { worksheets: { getItem: (n) => sheet(n), getItemOrNullObject: (n) => (sheets[n] ? sheet(n) : { isNullObject: true }) } };
}
test('add-in sync appends a new ticket right after the last ticket, not after blank rows', async () => {
  const H = BE.MASTER_COLS, blank = () => H.map(() => '');
  const raw = BE.FORM_COLS ? BE.FORM_COLS : H.slice(0, 24);
  const row = (cols, id) => cols.map((h) => (h === 'Id' ? id : h === 'Email' ? 'owner@example.com' : h === 'Completion time' ? '2026-10-05' : h === 'Status:' ? 'MHS - Submitted to Arietis' : ''));
  // Master has ticket 39 then blank rows whose only content is a stray label far below (simulates formatted/blank rows).
  const stray = blank(); stray[H.indexOf('Work Notes')] = 'x';
  const sheets = { Master: [H, row(H, 39), blank(), blank(), stray], 'Raw Data': [raw, row(raw, 39), row(raw, 40)], _SyncState: [['Id', 'Baseline JSON', 'Override Fields JSON', 'Conflicts JSON']], Settings: [['Setting', 'Value', '', 'Holiday dates']], 'Activity Log': [['Timestamp', 'User', 'Action', 'Ticket Id', 'Field', 'Old value', 'New value', 'Details']] };
  globalThis.Excel = { run: async (fn) => fn({ workbook: makeBook(sheets), sync: async () => {} }) };
  const r = await ExcelSource.sync('tester@example.com');
  assert.deepEqual(r.appended, [40]);
  assert.equal(sheets.Master[2][0], 40);
  const ids = sheets.Master.map((x) => x && x[0]).filter((v) => typeof v === 'number');
  assert.deepEqual(ids, [39, 40]);
});
