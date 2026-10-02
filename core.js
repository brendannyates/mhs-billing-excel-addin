export const EXTRA = ['Owner Email','Vendor Conversation ID','Submitted to Vendor At','Last Reply At','Closed At','Confirmation Due','Resolution Due','Escalation Flag','Sync Conflicts','Follow Up At','Reminder Flags'];
export const filled = v => v !== null && v !== undefined && String(v).trim() !== '';
export const same = (a,b) => String(a ?? '') === String(b ?? '');
export const closed = r => /^(Closed|Resolved)\b/.test(String(r['Status:']||''));
export function merge(master,raw,baseline={}) {
  const next={...master}, conflicts=[];
  for (const [key,value] of Object.entries(raw)) {
    if (!filled(value) || EXTRA.includes(key) || key==='Column1') continue;
    if (!filled(master[key])) next[key]=value;
    else if (!same(value,baseline[key]) && !same(value,master[key])) {
      if (Object.hasOwn(baseline,key) && same(master[key],baseline[key])) next[key]=value;
      else conflicts.push(key);
    }
  }
  return {next,conflicts};
}
export function localDay(v) {
  if (!filled(v)) return '';
  if (typeof v==='number') return new Date(Math.round((v-25569)*86400000)).toISOString().slice(0,10);
  if (/^\d{4}-\d{2}-\d{2}$/.test(String(v))) return String(v);
  const d=new Date(v); if(isNaN(d)) return '';
  return new Intl.DateTimeFormat('en-CA',{timeZone:'America/Los_Angeles',year:'numeric',month:'2-digit',day:'2-digit'}).format(d);
}
export function businessDue(start,days,holidays=[]) {
  const day=localDay(start); if(!day) return '';
  const d=new Date(day+'T12:00:00Z');let count=0;
  while(count<days){d.setUTCDate(d.getUTCDate()+1);const key=d.toISOString().slice(0,10);if(d.getUTCDay()!==0&&d.getUTCDay()!==6&&!holidays.includes(key))count++;}
  return d.toISOString().slice(0,10);
}
export function sla(row,holidays=[],today=localDay(new Date().toISOString())) {
  const start=row['Submitted to Vendor At'];
  const confirmation=businessDue(start,2,holidays),resolution=businessDue(start,5,holidays);
  const flags=[];
  if(!closed(row)){
    if(!filled(start)) flags.push('Vendor submission time needed');
    if(confirmation&&!filled(row['Date of first reply from Arietis:'])&&today>confirmation)flags.push('Receipt SLA overdue');
    if(resolution&&today>resolution)flags.push('Resolution SLA overdue');
    if(row['Urgency Level:']==='Critical')flags.push('Critical priority');
  }
  return {confirmation,resolution,flags};
}
export function businessAge(start,holidays=[],today=localDay(new Date().toISOString())) {
  const s=localDay(start);if(!s||s>today)return 0;const d=new Date(s+'T12:00:00Z');let age=0;
  while(d.toISOString().slice(0,10)<today){d.setUTCDate(d.getUTCDate()+1);const key=d.toISOString().slice(0,10);if(d.getUTCDay()!==0&&d.getUTCDay()!==6&&!holidays.includes(key))age++;}return age;
}
export const opened = r => r['Completion time'] || r['Start time'];
export function reminders(r,holidays=[],today=localDay(new Date().toISOString())) {
  const active=!closed(r),age=businessAge(opened(r),holidays,today);
  const followDates=[r['Follow Up At'],r['Date of Patient Outreach (if applicable):'],r['Last Reply At'],r['Date of first reply from Arietis:']].filter(filled).map(localDay).filter(Boolean).sort();
  const lastFollow=followDates.at(-1)||opened(r);
  return {age,awaiting:active&&!filled(r['Date of first reply from Arietis:']),needsFollowUp:active&&!filled(r['Date of first reply from Arietis:'])&&age>3,pastDue:active&&businessAge(lastFollow,holidays,today)>5,resolutionOverdue:active&&businessAge(r['Submitted to Vendor At']||opened(r),holidays,today)>6};
}
export function scopeTickets(rows,profile,{scope='assigned',mine=false,mrn='',clinic=''}={}) {
  return rows.filter(r=>(scope==='all'||(profile.clinics||[]).includes(r['Patient Clinic:']))&&(!clinic||r['Patient Clinic:']===clinic)&&(!mine||String(r['Owner Email']||r.Email).trim().toLowerCase()===String(profile.email||'').trim().toLowerCase())&&(!mrn||String(r['MRN:']).trim()===mrn.trim()));
}
const htmlEscape=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
export function clinicRecap(rows,clinic,mode='am',holidays=[],today=localDay(new Date().toISOString()),workbookUrl='') {
  const clinicRows=rows.filter(r=>r['Patient Clinic:']===clinic);
  const selected=mode==='pm'?clinicRows.filter(r=>closed(r)&&localDay(r['Closed At']||r['Date of Resolution:'])===today):clinicRows.filter(r=>!closed(r));
  const safeUrl=/^https:\/\//i.test(workbookUrl)?workbookUrl:'';
  const table=list=>list.length?'<table style="border-collapse:collapse;width:100%"><thead><tr>'+['Date Opened','MRN','Patient','Ticket Number','Reply Status','Resolution Status'].map(h=>'<th style="text-align:left;border:1px solid #c4d2ce;padding:8px">'+h+'</th>').join('')+'</tr></thead><tbody>'+list.map(r=>'<tr>'+[localDay(opened(r)),r['MRN:'],r['Patient:'],'#'+r.Id,filled(r['Date of first reply from Arietis:'])?'Confirmed '+localDay(r['Date of first reply from Arietis:']):'Awaiting reply',r['Status:']||'Unresolved'].map(v=>'<td style="border:1px solid #c4d2ce;padding:8px">'+htmlEscape(v)+'</td>').join('')+'</tr>').join('')+'</tbody></table>':'<p>No tickets in this section.</p>';
  const needs=selected.filter(r=>reminders(r,holidays,today).needsFollowUp),overdue=selected.filter(r=>reminders(r,holidays,today).resolutionOverdue);
  const intro=mode==='pm'?'Please see the below billing tickets resolved or closed today.':'Please see the below open billing tickets';
  return {subject:`${clinic} — ${mode==='pm'?'End-of-day billing recap':'AM open billing tickets'} — ${today}`,count:selected.length,html:`<div style="font:14px Arial;color:#203d42"><p>${intro}</p><h2>${htmlEscape(clinic)} · ${htmlEscape(today)}</h2>${table(selected)}${mode==='am'?'<h3>Needs follow-up — beyond 3 business days, receipt unconfirmed</h3>'+table(needs)+'<h3>Resolution Status Overdue — more than 6 business days</h3>'+table(overdue):''}${safeUrl?'<p><a href="'+htmlEscape(safeUrl)+'" style="background:#173f43;color:white;padding:10px 16px;text-decoration:none">Update tickets — open workbook and add-in</a></p>':'<p>Add the workbook URL in Settings to include an Update tickets link.</p>'}<p>SLA: 2 business days for receipt confirmation; 5 business days for resolution. Reminder thresholds include submission timing grace.</p></div>`};
}
