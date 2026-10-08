/**
 * Billing Digests — personalized status emails, read-only (never writes the workbook).
 * Run from Power Automate after the serialized worker has synced.
 * mode: "open" (start of day) | "close" (end of day) | "weekly". localNow: Pacific time, yyyy-MM-ddTHH:mm:ss.
 * Returns JSON text: { skip, reason, count, emails: [{ to, cc, subject, html, kind }] }.
 * Owner/leadership digests omit patient code + MRN unless Settings "Digest include patient IDs" = Yes.
 */
function main(workbook: ExcelScript.Workbook, mode: string = "open", localNow: string = ""): string {
  const now = toSerial(localNow);
  if (now === null) throw Error("Pass localNow as Pacific yyyy-MM-ddTHH:mm:ss.");
  const read = (name: string): Cell[][] => {
    const ws = workbook.getWorksheet(name);
    return ws ? (ws.getUsedRange().getValues() as Cell[][]) : [];
  };
  const master = read("Master");
  if (!master.length) throw Error("Master sheet not found. Initialize with the Excel add-in first.");
  const settingsRows = read("Settings");
  const settings = readSettings(settingsRows.slice(1).map((r) => [r[0], r[1]] as Cell[]));
  const hol = holidaySet(settingsRows.slice(1).map((r) => r[3]).filter((v) => /^\d{4}-\d{2}-\d{2}$/.test(String(v || ""))));
  const ctx: EvalCtx = { now: now, hol: hol, receiptDays: settings.receiptDays, resolutionDays: settings.resolutionDays, staleDays: settings.staleDays };
  if (mode !== "weekly" && !isBizDay(Math.floor(now), hol)) return JSON.stringify({ skip: true, reason: "Not a business day", count: 0, emails: [] });
  const tickets = loadTickets(master[0], master.slice(1), ctx);
  // Activity Log: Timestamp, User, Action, Ticket Id, Field, Old value, New value, Details
  const activity: ActivityRow[] = read("Activity Log").slice(1).map((r) => ({
    at: toSerial(r[0]) || 0, id: toNum(r[3]) || 0, field: norm(r[4]) || norm(r[2]), oldVal: norm(r[5]), newVal: norm(r[6]), by: norm(r[1]), source: norm(r[2]),
  }));
  let emails = buildDigests(tickets, activity, ctx, settings, mode);
  if (mode === "open" && settings.clinicDigest) {
    const ref = read("REF");
    const ix = buildIndex(ref.length ? ref[0] : []);
    const map = new Map<string, string>();
    ref.slice(1).forEach((r) => { const n = norm(cell(r, ix, "ClinicName")); const e = norm(cell(r, ix, "ClinicEmail")); if (n && e) map.set(n.toLowerCase(), e); });
    emails = emails.concat(buildClinicDigests(tickets, map, ctx, settings));
  }
  if (mode === "open" && settings.arietisDigest) {
    const f = buildArietisFollowup(tickets, settings, ctx, "MHS Billing");
    if (f) emails.push(f);
  }
  return JSON.stringify({ skip: false, reason: "", count: emails.length, emails: emails });
}
