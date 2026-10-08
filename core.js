// Sync merge policy (tested in tests/core.test.mjs). SLA clocks, actions and recaps live in src/rules.ts -> rules.js.
export const EXTRA = ['Owner Email','Vendor Conversation ID','Submitted to Vendor At','Last Reply At','Closed At','Confirmation Due','Resolution Due','Escalation Flag','Sync Conflicts','Follow Up At','Reminder Flags','CC','Last Updated By','Last Updated At','Work Notes','Ops Escalation','Ops Reason','Ops Escalated By','Ops Escalated At','Arietis Commitment','Follow-up Due','Review Notes','Last Reviewed At'];
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
export const opened = r => r['Completion time'] || r['Start time'];
export function scopeTickets(rows,profile,{scope='assigned',mine=false,mrn='',clinic=''}={}) {
  const clinics=profile.clinics||[],all=scope==='all'||!clinics.length;
  return rows.filter(r=>(all||clinics.includes(r['Patient Clinic:']))&&(!clinic||r['Patient Clinic:']===clinic)&&(!mine||String(r.Email||'').trim().toLowerCase()===String(profile.email||'').trim().toLowerCase())&&(!mrn||String(r['MRN:']).trim()===mrn.trim()));
}
