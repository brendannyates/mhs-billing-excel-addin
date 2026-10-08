/** Billing Clinic Recap — AM/PM clinic email templates, read-only, never sends email.
 *  localToday: Pacific YYYY-MM-DD; mode: am | pm; clinic: blank = every REF clinic; workbookUrl: optional (else Settings B6).
 *  Returns JSON text: { skip, clinicDigests: [{ clinic, email, subject, count, needsFollowUp, resolutionOverdue, html }] }. */
function main(workbook: ExcelScript.Workbook, localToday: string, mode: string = "am", clinic: string = "", workbookUrl: string = ""): string {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(localToday)) throw Error("Pass Pacific localToday YYYY-MM-DD");
  if (["am", "pm"].indexOf(mode) < 0) throw Error("Use am or pm");
  const mv = workbook.getWorksheet("Master").getUsedRange().getValues() as Cell[][];
  const ref = (workbook.getWorksheet("REF").getUsedRange().getValues() as Cell[][]).slice(1);
  const settingsRows = workbook.getWorksheet("Settings").getUsedRange().getValues() as Cell[][];
  const cfg = readSettings(settingsRows.slice(1).map((r) => [r[0], r[1]] as Cell[]));
  const hol = holidaySet(settingsRows.slice(1).map((r) => String(r[3] || "")).filter((v) => /^\d{4}-\d{2}-\d{2}$/.test(v)));
  const today = toSerial(localToday) as number;
  if (!isBizDay(today, hol)) return JSON.stringify({ skip: true, clinicDigests: [] });
  const ctx: EvalCtx = { now: today + (mode === "pm" ? 0.7 : 0.35), hol: hol, receiptDays: cfg.receiptDays, resolutionDays: cfg.resolutionDays, staleDays: cfg.staleDays };
  const tickets = loadTickets(mv[0], mv.slice(1), ctx);
  const url = workbookUrl || String((settingsRows[5] && settingsRows[5][1]) || "");
  const clinics = ref.map((r) => norm(r[0])).filter((c, i, a) => c && a.indexOf(c) === i && (!clinic || c === clinic));
  if (clinic && !clinics.length) throw Error("Clinic not found in REF");
  const clinicDigests = clinics.map((c) => {
    const r = buildClinicRecap(tickets, c, mode, ctx, url);
    const row = ref.find((x) => norm(x[0]) === c);
    return { clinic: c, email: row ? norm(row[1]) : "", subject: r.subject, count: r.count, needsFollowUp: r.needsFollowUp, resolutionOverdue: r.resolutionOverdue, html: r.html };
  }).filter((d) => d.count > 0);
  return JSON.stringify({ skip: false, clinicDigests: clinicDigests });
}
