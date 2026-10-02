export const EXTRA = ['Owner Email','Vendor Conversation ID','Submitted to Vendor At','Last Reply At','Closed At','Confirmation Due','Resolution Due','Escalation Flag','Sync Conflicts'];
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
