/**
 * DigestsDv — Power Pages / Dataverse edition of the digest builder (Billing Digests reads the workbook instead).
 * The flow lists rows from Dataverse and passes them in as JSON text; this script returns the
 * personalized emails. It doesn't read the workbook it runs in (any blank helper workbook works).
 * mode: "open" | "close" | "weekly". nowLocal: Pacific time as yyyy-MM-ddTHH:mm:ss.
 */
interface DigestOutDv {
  skip: boolean;
  reason: string;
  count: number;
  emails: Email[];
}
function main(workbook: ExcelScript.Workbook, mode: string, nowLocal: string, ticketsJson: string, activityJson: string, settingsJson: string, holidaysJson: string, refJson: string): DigestOutDv {
  const parsedNow = toSerial(nowLocal);
  if (parsedNow === null) throw Error("Pass nowLocal as Pacific yyyy-MM-ddTHH:mm:ss.");
  const now: number = parsedNow;
  const parse = (s: string): DvRecord[] => (s ? (JSON.parse(s) as DvRecord[]) : []);
  const settings = readSettings(parse(settingsJson).map((r) => [norm(r["mhs_name"] as string), norm(r["mhs_value"] as string)] as Cell[]));
  const hol = holidaySet(parse(holidaysJson).map((r) => norm(r["mhs_date"] as string).slice(0, 10)));
  const ctx: EvalCtx = { now: now, hol: hol, receiptDays: settings.receiptDays, resolutionDays: settings.resolutionDays, staleDays: settings.staleDays };
  if (mode !== "weekly" && !isBizDay(Math.floor(now), hol)) return { skip: true, reason: "Not a business day", count: 0, emails: [] };
  const t = dvTicketsToTable(parse(ticketsJson), pacificOffsetMin);
  const tickets = loadTickets(t.headers, t.rows, ctx);
  const activity = dvActivityRows(parse(activityJson), pacificOffsetMin);
  let emails = buildDigests(tickets, activity, ctx, settings, mode);
  if (mode === "open" && settings.clinicDigest) {
    const map = new Map<string, string>();
    for (const r of parse(refJson)) if (norm(r["mhs_category"] as string) === "ClinicName" && r["mhs_email"]) map.set(normKey(r["mhs_name"] as string), norm(r["mhs_email"] as string));
    emails = emails.concat(buildClinicDigests(tickets, map, ctx, settings));
  }
  if (mode === "open" && settings.arietisDigest) {
    const f = buildArietisFollowup(tickets, settings, ctx, "MHS Billing");
    if (f) emails.push(f);
  }
  return { skip: false, reason: "", count: emails.length, emails: emails };
}
