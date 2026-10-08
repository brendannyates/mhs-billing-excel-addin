// Tests for the shared rules engine (src/rules.ts -> rules.js): business days, SLA clocks from submission,
// needs-action list, ops escalation, vendor-safe export, derived Master columns, clinic recaps, digests.
import test from 'node:test'; import assert from 'node:assert/strict'; import fs from 'node:fs'; import vm from 'node:vm';
const ctxVm = vm.createContext({}); vm.runInContext(fs.readFileSync(new URL('../rules.js', import.meta.url), 'utf8'), ctxVm);
const BE = ctxVm.BE, S = (v) => BE.toSerial(v);
const HOL = BE.holidaySet(['2026-09-07', '2026-10-05', '2026-11-26', '2026-11-27']);
const H = BE.MASTER_COLS;
const row = (o) => { const r = H.map(() => ''); for (const k in o) r[H.indexOf(k)] = o[k]; return r; };
const ctx = (now) => ({ now: S(now), hol: HOL, receiptDays: 2, resolutionDays: 5, staleDays: 3 });
const one = (o, now) => BE.loadTickets(H, [row({ Id: 1, 'Completion time': '2026-10-02T10:00', 'Status:': BE.STATUS.submitted, ...o })], ctx(now))[0];

test('business days skip weekends and listed holidays', () => {
  assert.equal(BE.fmtDate(BE.addBizDays(S('2026-10-08T16:00'), 2, HOL)), '10/12/26'); // Thu 4pm -> Mon
  assert.equal(BE.fmtDate(BE.addBizDays(S('2026-10-10'), 2, HOL)), '10/13/26');       // Sat -> Tue
  assert.equal(BE.fmtDate(BE.addBizDays(S('2026-09-04'), 2, HOL)), '9/9/26');         // Labor Day skipped
  assert.equal(BE.fmtDate(BE.addBizDays(S('2026-11-25'), 2, HOL)), '12/1/26');        // Thanksgiving + Fri skipped
  assert.equal(BE.fmtDate(BE.addBizDays(S('2026-10-02'), 2, HOL)), '10/7/26');        // Fri, Mon holiday -> Wed
});

test('receipt and resolution clocks start at submission; due day is not overdue', () => {
  const t = one({}, '2026-10-07T15:00');
  assert.equal(BE.serialToIsoDate(t.receipt.due), '2026-10-07');
  assert.equal(t.receipt.state, 'due-today');
  assert.equal(BE.serialToIsoDate(t.resolution.due), '2026-10-12');
  const late = one({}, '2026-10-08T09:00');
  assert.equal(late.receipt.state, 'breached');
  assert.equal(late.actions[0].code, 'RCPT_OVERDUE');
  assert.equal(late.actions[0].sev, 1);
  assert.equal(late.actions[0].days, 1);
  assert.equal(late.actions[0].label, 'No receipt from Arietis');
  assert.match(late.actions[0].text, /1 day overdue/);
  assert.equal(BE.slaText(late.receipt), '1 day overdue');
});

test('no "BD" abbreviation in action text or SLA text', () => {
  const t = one({ 'Last Updated At': '2026-10-02' }, '2026-10-20');
  for (const a of t.actions) assert.ok(!/\bBD\b/.test(a.text), a.text);
  assert.ok(!/\bBD\b/.test(BE.slaText(t.resolution)));
});

test('first reply stops the receipt clock; confirmed status without a date counts as met', () => {
  assert.equal(one({ 'Date of first reply from Arietis:': '2026-10-06T21:30:00Z' }, '2026-10-09').receipt.state, 'met');
  const nodate = one({ 'Status:': BE.STATUS.confirmed }, '2026-10-09');
  assert.equal(nodate.receipt.state, 'met-nodate');
  assert.ok(nodate.actions.some((a) => a.code === 'RCPT_DATE'));
});

test('ISO timestamps with Z convert to Pacific wall-clock time', () => {
  const s = S('2026-10-06T21:30:00Z'); // 2:30 pm PDT
  assert.equal(BE.fmtDateTime(s), '10/6 2:30pm');
});

test('closed tickets: resolution measured at Date of Resolution, missing closure fields flagged', () => {
  const t = one({ 'Status:': BE.STATUS.resolved, 'Date of Resolution:': '2026-10-13' }, '2026-10-20');
  assert.equal(t.resolution.state, 'late');
  assert.ok(t.actions.some((a) => a.code === 'CLOSE_INCOMPLETE'));
});

test('sync conflicts surface as an owner action', () => {
  const t = one({ 'Sync Conflicts': 'Notes' }, '2026-10-05');
  assert.ok(t.actions.some((a) => a.code === 'SYNC_CONFLICT' && a.who === 'owner'));
});

test('ops queue: Critical, service recovery and manual flags; Cleared wins; overdue follow-up is severity 1', () => {
  const ts = BE.loadTickets(H, [
    row({ Id: 1, 'Completion time': '2026-10-01', 'Urgency Level:': '*Critical*' }),
    row({ Id: 2, 'Completion time': '2026-10-01', 'Patient Feedback/Service Recovery Flag:': 'Yes' }),
    row({ Id: 3, 'Completion time': '2026-10-01', 'Ops Escalation': 'Yes', 'Ops Reason': 'Repeat complaint', 'Follow-up Due': '2026-10-05' }),
    row({ Id: 4, 'Completion time': '2026-10-01', 'Urgency Level:': 'Critical', 'Ops Escalation': 'Cleared' }),
    row({ Id: 5, 'Completion time': '2026-10-01' }),
  ], ctx('2026-10-07'));
  assert.deepEqual([...BE.opsQueue(ts).map((t) => t.id)].sort(), [1, 2, 3]);
  assert.ok(ts.find((t) => t.id === 3).actions.some((a) => a.code === 'OPS_FU_OVERDUE' && a.sev === 1));
  const exported = JSON.stringify(BE.arietisExportRows(ts.filter((t) => t.opsEscalated)));
  assert.ok(!exported.includes('Repeat complaint'), 'internal ops reason must never be exported');
});

test('derived Master columns: due dates as ISO days, overdue items in Escalation Flag', () => {
  const t = one({ 'Urgency Level:': 'Critical' }, '2026-10-13');
  const d = Object.fromEntries(BE.derivedColumns(t));
  assert.equal(d['Confirmation Due'], '2026-10-07');
  assert.equal(d['Resolution Due'], '2026-10-12');
  assert.match(d['Escalation Flag'], /No receipt from Arietis/);
  assert.match(d['Escalation Flag'], /Critical priority/);
});

test('clinic AM recap: patient fields, both overdue sections, safe workbook link', () => {
  const ts = BE.loadTickets(H, [row({ Id: 7, 'Completion time': '2026-10-02', 'Patient Clinic:': 'San Rafael - CA', 'MRN:': '00123', 'Patient:': 'TEST', 'Status:': BE.STATUS.submitted })], ctx('2026-10-14'));
  const r = BE.buildClinicRecap(ts, 'San Rafael - CA', 'am', ctx('2026-10-14'), 'https://example.com/wb');
  assert.equal(r.count, 1); assert.equal(r.needsFollowUp, 1); assert.equal(r.resolutionOverdue, 1);
  for (const h of ['Date Opened', 'MRN', 'Patient', 'Ticket Number', 'Reply Status', 'Resolution Status', '00123', 'https://example.com/wb']) assert.ok(r.html.includes(h), h);
  const bad = BE.buildClinicRecap(BE.loadTickets(H, [row({ Id: 8, 'Completion time': '2026-10-02', 'Patient Clinic:': 'X', 'Patient:': '<script>alert(1)</script>' })], ctx('2026-10-14')), 'X', 'am', ctx('2026-10-14'), 'javascript:alert(1)');
  assert.ok(!bad.html.includes('<script>')); assert.ok(!bad.html.includes('javascript:'));
});

test('clinic PM recap uses the Pacific closure day', () => {
  const ts = BE.loadTickets(H, [
    row({ Id: 1, 'Completion time': '2026-10-01', 'Patient Clinic:': 'C', 'Status:': BE.STATUS.resolved, 'Closed At': '2026-10-03T01:00:00Z' }), // 6pm Oct 2 PDT
    row({ Id: 2, 'Completion time': '2026-10-01', 'Patient Clinic:': 'C', 'Status:': BE.STATUS.closedNoAction, 'Closed At': '2026-10-03T20:00:00Z' }),
  ], ctx('2026-10-02T17:00'));
  assert.equal(BE.buildClinicRecap(ts, 'C', 'pm', ctx('2026-10-02T17:00'), '').count, 1);
});

test('owner digests omit patient code and MRN unless Settings allow them', () => {
  const ts = BE.loadTickets(H, [row({ Id: 9, 'Completion time': '2026-10-01', Email: 'owner@example.com', 'MRN:': '00999', 'Patient:': 'ZZTOP' })], ctx('2026-10-07T08:00'));
  const off = JSON.stringify(BE.buildDigests(ts, [], ctx('2026-10-07T08:00'), BE.readSettings([]), 'open'));
  assert.ok(off.includes('owner@example.com')); assert.ok(!off.includes('00999')); assert.ok(!off.includes('ZZTOP'));
  const on = JSON.stringify(BE.buildDigests(ts, [], ctx('2026-10-07T08:00'), BE.readSettings([['Digest include patient IDs', 'Yes']]), 'open'));
  assert.ok(on.includes('00999'));
});

test('settings read from the workbook Settings sheet labels', () => {
  const s = BE.readSettings([['Receipt SLA business days', 3], ['Resolution SLA business days', 7], ['Leadership recipients', 'a@x.com; b@x.com']]);
  assert.equal(s.receiptDays, 3); assert.equal(s.resolutionDays, 7); assert.deepEqual([...s.leadership], ['a@x.com', 'b@x.com']);
});
