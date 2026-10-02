/** Generate clinic email templates inside Microsoft 365; never sends email. */
type Cell = string | number | boolean;
type Ticket = {[key: string]: Cell};
function escapeHtml(v: Cell | undefined): string {return String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));}
function filled(v: Cell | undefined): boolean {return v!==undefined&&String(v).trim()!=='';}
function closed(t: Ticket): boolean {return /^(Closed|Resolved)\b/.test(String(t['Status:']||''));}
function day(v: Cell): string {
  if(!filled(v))return '';if(typeof v==='number')return new Date(Math.round((v-25569)*86400000)).toISOString().slice(0,10);
  if(/^\d{4}-\d{2}-\d{2}$/.test(String(v)))return String(v);const d=new Date(String(v));if(isNaN(d.getTime()))return '';
  return new Intl.DateTimeFormat('en-CA',{timeZone:'America/Los_Angeles',year:'numeric',month:'2-digit',day:'2-digit'}).format(d);
}
function age(v: Cell, today: string, holidays: string[]): number {const start=day(v);if(!start||start>today)return 0;const d=new Date(start+'T12:00:00Z');let n=0;while(d.toISOString().slice(0,10)<today){d.setUTCDate(d.getUTCDate()+1);if(d.getUTCDay()!==0&&d.getUTCDay()!==6&&!holidays.includes(d.toISOString().slice(0,10)))n++;}return n;}
function opened(t: Ticket): Cell {return t['Completion time']||t['Start time']||'';}
function table(rows: Ticket[]): string {
  if(!rows.length)return '<p>No tickets in this section.</p>';
  return '<table style="border-collapse:collapse;width:100%"><thead><tr>'+['Date Opened','MRN','Patient','Ticket Number','Reply Status','Resolution Status'].map(h=>'<th style="text-align:left;border:1px solid #c4d2ce;padding:8px">'+h+'</th>').join('')+'</tr></thead><tbody>'+rows.map(t=>'<tr>'+[day(opened(t)),t['MRN:'],t['Patient:'],'#'+t.Id,filled(t['Date of first reply from Arietis:'])?'Confirmed '+day(t['Date of first reply from Arietis:']):'Awaiting reply',t['Status:']||'Unresolved'].map(v=>'<td style="border:1px solid #c4d2ce;padding:8px">'+escapeHtml(v)+'</td>').join('')+'</tr>').join('')+'</tbody></table>';
}
function main(workbook: ExcelScript.Workbook, localToday: string, mode: string='am', clinic: string='', workbookUrl: string=''): string {
  if(!/^\d{4}-\d{2}-\d{2}$/.test(localToday))throw Error('Pass Pacific localToday YYYY-MM-DD');if(!['am','pm'].includes(mode))throw Error('Use am or pm');
  const mv=workbook.getWorksheet('Master').getUsedRange().getValues(),headers=mv[0].map(String);
  const tickets=mv.slice(1).filter(r=>filled(r[0])).map(r=>{const t: Ticket={};headers.forEach((h,i)=>t[h]=r[i]??'');return t;});
  const ref=workbook.getWorksheet('REF').getUsedRange().getValues().slice(1),settings=workbook.getWorksheet('Settings').getUsedRange().getValues();
  const holidays=settings.slice(1).map(r=>String(r[3]||'')).filter(v=>/^\d{4}-\d{2}-\d{2}$/.test(v));
  if(holidays.includes(localToday)||[0,6].includes(new Date(localToday+'T12:00:00Z').getUTCDay()))return JSON.stringify({skip:true,clinicDigests:[]});
  const url=workbookUrl||String(settings[5]?.[1]||'');const safeUrl=/^https:\/\//i.test(url)?url:'';
  const clinics=[...new Set(ref.map(r=>String(r[0]||'')).filter(c=>c&&(!clinic||c===clinic)))];
  if(clinic&&!clinics.length)throw Error('Clinic not found in REF');
  const clinicDigests=clinics.map(c=>{
    const selected=tickets.filter(t=>t['Patient Clinic:']===c&&(mode==='am'?!closed(t):closed(t)&&day(t['Closed At']||t['Date of Resolution:'])===localToday));
    const needs=selected.filter(t=>!closed(t)&&!filled(t['Date of first reply from Arietis:'])&&age(opened(t),localToday,holidays)>3);
    const overdue=selected.filter(t=>!closed(t)&&age(t['Submitted to Vendor At']||opened(t),localToday,holidays)>6);
    const intro=mode==='am'?'Please see the below open billing tickets':'Please see the below billing tickets resolved or closed today.';
    const html='<div style="font:14px Arial;color:#203d42"><p>'+intro+'</p><h2>'+escapeHtml(c)+' · '+localToday+'</h2>'+table(selected)+(mode==='am'?'<h3>Needs follow-up — beyond 3 business days, receipt unconfirmed</h3>'+table(needs)+'<h3>Resolution Status Overdue — more than 6 business days</h3>'+table(overdue):'')+(safeUrl?'<p><a href="'+escapeHtml(safeUrl)+'">Update tickets — open workbook and add-in</a></p>':'')+'<p>SLA: 2 business days for receipt confirmation; 5 business days for resolution. Reminder thresholds include submission timing grace.</p></div>';
    return {clinic:c,email:String(ref.find(r=>r[0]===c)?.[1]||''),subject:c+' — '+(mode==='am'?'AM open billing tickets':'End-of-day billing recap')+' — '+localToday,count:selected.length,needsFollowUp:needs.length,resolutionOverdue:overdue.length,html};
  }).filter(d=>d.count>0);
  return JSON.stringify({skip:false,clinicDigests});
}
