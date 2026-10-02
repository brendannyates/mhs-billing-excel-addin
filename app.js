import {EXTRA,filled,merge,closed,sla,reminders,scopeTickets,clinicRecap,localDay,opened} from './core.js';
const $=s=>document.querySelector(s), esc=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const FORM='https://forms.cloud.microsoft/Pages/ResponsePage.aspx?id=8Sq2y8CkOEWdeGKInLNdYw1_rrxVxaFImLkzEWnt0GFUQUtJVFg2VzU5UEowM01JMUgzUUpKMTZPVi4u';
let rows=[],headers=[],refs=[],holidays=[],current='clinic',filter='',editing=null,busy=false,scope='assigned',mine=false,mrn='',statusSection='unresolved',workbookUrl='',ready=false;
let profile={name:'',email:'',clinics:[],role:''};try{profile={...profile,...JSON.parse(localStorage.getItem('mhsBillingProfile')||'{}')};}catch{}
const options={'Department Submitting Ticket:':2,'Patient Clinic:':3,'Urgency Level:':5,'Source of Inquiry:':6,'Task Type:':7,'Status:':10,'Was false verification of balance/charge given by Arietis?':14,'Patient Feedback/Service Recovery Flag:':15,'Source of Error:':16,'Outcome:':17};
function objects(v){return v.slice(1).filter(r=>filled(r[0])).map((r,index)=>({...Object.fromEntries(v[0].map((h,i)=>[h,r[i]??''])),_row:index}));}
async function run(fn){if(busy)return;busy=true;try{await fn();$('#message').textContent=`Updated ${new Date().toLocaleTimeString()} · Auto-sync every 30 seconds while open`;}catch(e){$('#message').textContent=e.message;}finally{busy=false;}}
async function initialize(){await Excel.run(async ctx=>{
  const master=ctx.workbook.worksheets.getItem('Master');const range=master.getUsedRange();range.load('values');await ctx.sync();
  const h=range.values[0].map(String);for(const name of EXTRA)if(!h.includes(name)){master.getCell(0,h.length).values=[[name]];h.push(name);}
  const tabs=ctx.workbook.worksheets;tabs.load('items/name');await ctx.sync();
  if(!tabs.items.some(x=>x.name==='_SyncState')){const s=tabs.add('_SyncState');s.getRange('A1:D1').values=[['Id','Baseline JSON','Override Fields JSON','Conflicts JSON']];s.visibility='Hidden';}
  if(!tabs.items.some(x=>x.name==='Settings')){const s=tabs.add('Settings');s.getRange('A1:B5').values=[['Setting','Value'],['Time zone','Pacific Standard Time'],['Business open','09:00'],['Business close','17:00'],['Weekly recap day','Friday']];s.getRange('D1').values=[['Holiday dates (YYYY-MM-DD)']];s.getRange('D2').values=[['']];}
  master.tables.load('items/name');await ctx.sync();if(!master.tables.items.length){const t=master.tables.add(master.getRangeByIndexes(0,0,range.values.length,h.length),true);t.name='BillingMaster';}
  else {master.tables.items[0].resize(master.getRangeByIndexes(0,0,range.values.length,h.length));}
  await ctx.sync();
});await refresh();}
async function refresh(){await Excel.run(async ctx=>{
  const m=ctx.workbook.worksheets.getItem('Master').getUsedRange();const r=ctx.workbook.worksheets.getItem('REF').getUsedRange();const s=ctx.workbook.worksheets.getItemOrNullObject('Settings');s.load('isNullObject');m.load('values');r.load('values');await ctx.sync();
  headers=m.values[0].map(String);rows=objects(m.values);refs=r.values.slice(1);holidays=[];
  if(!s.isNullObject){const v=s.getUsedRange();v.load('values');await ctx.sync();holidays=v.values.slice(1).map(x=>String(x[3]||'')).filter(x=>/^\d{4}-\d{2}-\d{2}$/.test(x));workbookUrl=v.values[5]?.[1]||'';}
});render();}
async function sync(){await Excel.run(async ctx=>{
  const ws=ctx.workbook.worksheets.getItem('Master'),raw=ctx.workbook.worksheets.getItem('Raw Data'),state=ctx.workbook.worksheets.getItem('_SyncState');
  const mr=ws.getUsedRange(),rr=raw.getUsedRange(),sr=state.getUsedRange();mr.load('values');rr.load('values');sr.load('values');await ctx.sync();
  const mh=mr.values[0].map(String),rh=rr.values[0].map(String),index=new Map();
  mr.values.slice(1).forEach((r,i)=>{if(filled(r[0])){if(index.has(String(r[0])))throw Error('Duplicate Id in Master; sync stopped.');index.set(String(r[0]),i+1);}});
  const states=new Map(sr.values.slice(1).filter(r=>filled(r[0])).map(r=>[String(r[0]),r]));let append=mr.values.length;
  const seen=new Set();for(const rawRow of rr.values.slice(1)){
    if(!filled(rawRow[0]))continue;const id=String(rawRow[0]);if(seen.has(id))throw Error('Duplicate Id in Raw Data; sync stopped.');seen.add(id);
  }
  for(const rawRow of rr.values.slice(1)){
    if(!filled(rawRow[0]))continue;const id=String(rawRow[0]),r=Object.fromEntries(rh.map((h,i)=>[h,rawRow[i]??'']));
    const pos=index.get(id),m=pos===undefined?{}:Object.fromEntries(mh.map((h,i)=>[h,mr.values[pos][i]??'']));
    const old=states.get(id),baseline=old?JSON.parse(old[1]||'{}'):{},overrides=old?JSON.parse(old[2]||'[]'):[];
    const out=merge(m,r,baseline);for(const key of overrides)out.next[key]=m[key]??'';
    if(closed(m))out.next['Status:']=m['Status:'];
    if(filled(m['Date of first reply from Arietis:']))out.next['Date of first reply from Arietis:']=m['Date of first reply from Arietis:'];
    const pending=new Set(old?JSON.parse(old[3]||'[]'):[]);out.conflicts.forEach(k=>pending.add(k));
    if(!filled(out.next['Owner Email']))out.next['Owner Email']=out.next.Email||'';
    if(closed(out.next)&&!closed(m)&&!filled(out.next['Closed At'])){out.next['Status:']=m['Status:']||'MHS - Submitted to Arietis';pending.add('Status: (close in add-in after inbox reconciliation)');}
    const grace=reminders(out.next,holidays);out.next['Reminder Flags']=[grace.needsFollowUp?'Needs follow-up beyond 3 business days':'',grace.pastDue?'No follow-up beyond 5 business days':'',grace.resolutionOverdue?'Resolution overdue beyond 6 business days':''].filter(Boolean).join('; ');
    const flags=sla(out.next,holidays);out.next['Confirmation Due']=flags.confirmation;out.next['Resolution Due']=flags.resolution;out.next['Escalation Flag']=flags.flags.join('; ');out.next['Sync Conflicts']=[...pending].join('; ');
    if(pos===undefined){ws.getRangeByIndexes(append,0,1,mh.length).values=[mh.map(h=>out.next[h]??'')];index.set(id,append++);}
    else for(let j=0;j<mh.length;j++)if(String(out.next[mh[j]]??'')!==String(mr.values[pos][j]??''))ws.getCell(pos,j).values=[[out.next[mh[j]]??'']];
    states.set(id,[id,JSON.stringify(r),JSON.stringify(overrides),JSON.stringify([...pending])]);
  }
  if(states.size)state.getRangeByIndexes(1,0,states.size,4).values=[...states.values()];
  ws.tables.load('items/name');await ctx.sync();if(ws.tables.items.length)ws.tables.items[0].resize(ws.getRangeByIndexes(0,0,append,mh.length));await ctx.sync();
});await refresh();}
function filtered(){return scopeTickets(rows,profile,{scope,mine,mrn,clinic:filter});}
function clinicOptions(){return [...new Set([...refs.map(r=>r[0]),...rows.map(r=>r['Patient Clinic:'])].filter(filled))];}
function ticketCard(r){const x=reminders(r,holidays),formal=sla(r,holidays);return `<article class="card"><strong>Ticket #${esc(r.Id)} · ${esc(r['Patient Clinic:'])}</strong><p class="meta">${esc(r['Status:']||'Unresolved')} · ${esc(r['Urgency Level:'])}</p><p>MRN: ${esc(r['MRN:'])} · Patient: ${esc(r['Patient:'])}</p><p>Opened: ${esc(localDay(opened(r))||'—')} · ${x.age} business days<br>Receipt SLA: ${esc(formal.confirmation||'Vendor submission date needed')} · Resolution SLA: ${esc(formal.resolution||'—')}</p><p class="flag">${esc([x.needsFollowUp?'Needs follow-up: receipt unconfirmed beyond 3 business days':'',x.pastDue?'Past due: no follow-up within 5 business days':'',x.resolutionOverdue?'Resolution overdue: more than 6 business days':'',r['Sync Conflicts']?'Review sync conflicts: '+r['Sync Conflicts']:''].filter(Boolean).join(' · '))}</p><button data-id="${esc(r.Id)}">Update status</button></article>`;}
function render(){document.querySelectorAll('[data-view]').forEach(b=>b.classList.toggle('active',b.dataset.view===current));
  if(current==='form'){$('#view').innerHTML=`<h2>Submit New Ticket</h2><p><a href="${FORM}" target="_blank" rel="noopener noreferrer">Open form in a new window</a></p><iframe title="Billing escalation submission form" src="${FORM}&embed=true" width="640" height="480" frameborder="0" marginwidth="0" marginheight="0" style="border:none;max-width:100%;max-height:100vh" allowfullscreen></iframe>`;return;}
  if(current==='settings'){
    $('#view').innerHTML=`<h2>Your settings</h2><form id="profileForm"><label>Name<input name="name" value="${esc(profile.name)}" required></label><label>Email<input type="email" name="email" value="${esc(profile.email)}" required></label><label>Role<input name="role" value="${esc(profile.role)}" placeholder="Your clinic role" required></label><fieldset><legend>Assigned clinics</legend>${clinicOptions().map(c=>`<label class="check"><input type="checkbox" name="clinics" value="${esc(c)}" ${profile.clinics.includes(c)?'checked':''}>${esc(c)}</label>`).join('')}</fieldset><label>Cloud workbook URL (shared recap setting)<input type="url" name="workbookUrl" value="${esc(workbookUrl)}" placeholder="https://…"></label><button>Save settings</button></form><p>Profile settings are saved for this browser/device. Role labels your profile; workbook permissions control access. Edit organization holiday dates in Settings column D.</p><p>Formal SLA: 2 business days to confirm receipt and 5 to resolve. Follow-up reminders: beyond 3 business days from entry with no first reply. Resolution reminders: more than 6 business days. Past Due: no documented follow-up for more than 5 business days.</p>`;
    $('#profileForm').onsubmit=e=>{e.preventDefault();const data=new FormData(e.target),clinics=data.getAll('clinics');if(!clinics.length){$('#message').textContent='Select at least one assigned clinic.';return;}const url=String(data.get('workbookUrl')||'').trim();if(url&&!/^https:\/\//i.test(url)){$('#message').textContent='Use an HTTPS workbook URL.';return;}run(async()=>{if(url!==workbookUrl)await Excel.run(async ctx=>{ctx.workbook.worksheets.getItem('Settings').getRange('A6:B6').values=[['Workbook URL',url]];await ctx.sync();});const next={name:String(data.get('name')),email:String(data.get('email')).trim(),role:String(data.get('role')),clinics};localStorage.setItem('mhsBillingProfile',JSON.stringify(next));profile=next;workbookUrl=url;scope='assigned';mine=false;filter='';mrn='';current='clinic';render();});};return;
  }
  if(current==='recap'){renderRecap();return;}
  const a=filtered(),active=a.filter(r=>!closed(r)),awaiting=active.filter(r=>reminders(r,holidays).awaiting),past=active.filter(r=>reminders(r,holidays).pastDue),follow=active.filter(r=>reminders(r,holidays).needsFollowUp),overdue=active.filter(r=>reminders(r,holidays).resolutionOverdue);
  const sections={unresolved:active,awaiting,past,follow,overdue,all:a};
  const labels={unresolved:'Unresolved',awaiting:'Awaiting Reply',past:'Past Due',follow:'Needs Follow-up',overdue:'Resolution Overdue',all:'All statuses'};
  $('#view').innerHTML=`<h2>Clinic Dashboard</h2><p class="meta">${esc(profile.name||'Set up your profile in Settings')} ${profile.role?'· '+esc(profile.role):''}</p>${!profile.clinics.length?'<p class="notice">Choose your assigned clinics in Settings to populate this dashboard.</p>':''}<div class="filters"><label>Ticket scope<select id="scope"><option value="assigned" ${scope==='assigned'?'selected':''}>My assigned clinics</option><option value="all" ${scope==='all'?'selected':''}>All tickets / all clinics</option></select></label><label>Clinic<select id="clinicFilter"><option value="">${scope==='assigned'?'All assigned clinics':'All clinics'}</option>${clinicOptions().filter(c=>scope==='all'||profile.clinics.includes(c)).map(c=>`<option value="${esc(c)}" ${filter===c?'selected':''}>${esc(c)}</option>`).join('')}</select></label><label>Exact MRN<input id="mrnFilter" value="${esc(mrn)}" placeholder="Search within selected scope"></label><label class="check"><input type="checkbox" id="mine" ${mine?'checked':''}>My Tickets</label></div><div class="metrics"><div class="card">Open tickets<strong>${active.length}</strong></div><div class="card">Awaiting reply<strong>${awaiting.length}</strong></div><div class="card">Needs follow-up<strong>${follow.length}</strong></div><div class="card">Resolution overdue<strong>${overdue.length}</strong></div></div><div class="status-pills">${Object.entries(labels).map(([key,label])=>`<button data-section="${key}" class="${statusSection===key?'active':''}">${label} (${sections[key].length})</button>`).join('')}</div><h3>${labels[statusSection]}</h3>${sections[statusSection].length?sections[statusSection].map(ticketCard).join(''):'<p class="empty">No matching tickets.</p>'}`;
  $('#scope').onchange=e=>{scope=e.target.value;filter='';render();};$('#clinicFilter').onchange=e=>{filter=e.target.value;render();};$('#mrnFilter').onchange=e=>{mrn=e.target.value;render();};$('#mine').onchange=e=>{mine=e.target.checked;render();};document.querySelectorAll('[data-section]').forEach(b=>b.onclick=()=>{statusSection=b.dataset.section;render();});document.querySelectorAll('[data-id]').forEach(b=>b.onclick=()=>edit(b.dataset.id));
}
let recapClinic='',recapMode='am';
function renderRecap(){const clinics=clinicOptions();if(!clinics.includes(recapClinic))recapClinic=profile.clinics[0]||clinics[0]||'';const recap=clinicRecap(rows,recapClinic,recapMode,holidays,undefined,workbookUrl);const recipient=refs.find(r=>r[0]===recapClinic)?.[1]||'';
  $('#view').innerHTML=`<h2>Clinic email recap</h2><label>Clinic<select id="recapClinic">${clinics.map(c=>`<option ${c===recapClinic?'selected':''}>${esc(c)}</option>`).join('')}</select></label><label>Recap<select id="recapMode"><option value="am" ${recapMode==='am'?'selected':''}>AM — all open tickets</option><option value="pm" ${recapMode==='pm'?'selected':''}>End of day — resolved / closed today</option></select></label><p>To: ${esc(recipient||'Clinic email not configured')}<br>Subject: ${esc(recap.subject)}</p><button id="copyRecap">Copy email template</button><div class="notice">Preview only. This does not send email.</div><section id="recapPreview" class="card">${recap.html}</section><label>HTML template<textarea id="recapHtml" readonly>${esc(recap.html)}</textarea></label>`;
  $('#recapClinic').onchange=e=>{recapClinic=e.target.value;renderRecap();};$('#recapMode').onchange=e=>{recapMode=e.target.value;renderRecap();};$('#copyRecap').onclick=async()=>{try{await navigator.clipboard.write([new ClipboardItem({'text/html':new Blob([recap.html],{type:'text/html'}),'text/plain':new Blob([$('#recapPreview').innerText],{type:'text/plain'})})]);$('#message').textContent='Email template copied. Paste into your Microsoft 365 email.';}catch{$('#recapHtml').select();$('#message').textContent='Select/copy the HTML template below; clipboard access is unavailable.';}};
}

function edit(id){editing=rows.find(r=>String(r.Id)===id);const fields=headers.filter(h=>!['Id','Start time','Completion time','Confirmation Due','Resolution Due','Escalation Flag','Reminder Flags','Sync Conflicts','Column1','Date of first reply from Arietis:','Last Reply At','Closed At'].includes(h));$('#fields').innerHTML=fields.map(h=>{const v=editing[h]??'';const opts=options[h]===undefined?null:[...new Set(refs.map(r=>r[options[h]]).filter(x=>filled(x)&&x!=='[Free Text]'))];let control;if(opts&&h!=='Task Type:'){const all=[...new Set(['',v,...opts])];control=`<select name="${esc(h)}">${all.map(x=>`<option ${String(x)===String(v)?'selected':''}>${esc(x)}</option>`).join('')}</select>`;}else control=`<textarea name="${esc(h)}">${esc(v)}</textarea>`;return `<label>${esc(h)}${control}</label>`;}).join('');$('#editor').showModal();}
async function save(){const fd=new FormData($('#editForm'));await Excel.run(async ctx=>{
  const ws=ctx.workbook.worksheets.getItem('Master'),range=ws.getUsedRange(),state=ctx.workbook.worksheets.getItem('_SyncState'),sr=state.getUsedRange();range.load('values');sr.load('values');await ctx.sync();
  const h=range.values[0].map(String),pos=range.values.findIndex((r,i)=>i>0&&String(r[0])===String(editing.Id));if(pos<0)throw Error('Ticket missing; refresh.');
  const changes=[...fd].filter(([key,v])=>String(v)!==String(editing[key]??''));for(const [key] of changes)if(String(range.values[pos][h.indexOf(key)]??'')!==String(editing[key]??''))throw Error('Another editor changed this field; refresh before saving.');
  const newStatus=fd.get('Status:');const closing=/^(Closed|Resolved)\b/.test(String(newStatus));if(closed(editing)&&newStatus!==editing['Status:'])throw Error('Reopening needs an explicit reviewed workflow.');
  if(closing&&!closed(editing)&&!filled(fd.get('Date of Resolution:')))throw Error('Enter resolution date before closing.');
  for(const [key,v] of changes)ws.getCell(pos,h.indexOf(key)).values=[[v]];
  if(closing&&!closed(editing))ws.getCell(pos,h.indexOf('Closed At')).values=[[new Date().toISOString()]];
  const sp=sr.values.findIndex((r,i)=>i>0&&String(r[0])===String(editing.Id));const old=sp<0?[]:JSON.parse(sr.values[sp][2]||'[]');const override=[...new Set([...old,...changes.map(x=>x[0])])];
  const stateRow=sp<0?[String(editing.Id),'{}',JSON.stringify(override),'[]']:[...sr.values[sp]];stateRow[2]=JSON.stringify(override);stateRow[3]='[]';state.getRangeByIndexes(sp<0?sr.values.length:sp,0,1,4).values=[stateRow];ws.getCell(pos,h.indexOf('Sync Conflicts')).values=[['']];await ctx.sync();
});$('#editor').close();await refresh();}
function bindUI(){
 $('#setup').onclick=()=>run(initialize);$('#refresh').onclick=()=>run(refresh);$('#sync').onclick=()=>run(sync);
 document.querySelectorAll('[data-view]').forEach(b=>b.onclick=()=>{current=b.dataset.view;render();});
 $('#cancel').onclick=()=>$('#editor').close();$('#editForm').onsubmit=e=>{e.preventDefault();run(save);};
}
render();document.querySelectorAll('[data-view]').forEach(b=>b.onclick=()=>{current=b.dataset.view;render();});
Office.onReady(info=>{if(info.host!==Office.HostType.Excel){render();document.querySelectorAll('[data-view]').forEach(b=>b.onclick=()=>{current=b.dataset.view;render();});return;}
 ready=true;bindUI();run(initialize);
 setInterval(()=>{if(ready&&!busy&&!$('#editor').open&&current!=='settings'&&!document.activeElement?.matches('input,textarea,select'))run(sync);},30000);
});
