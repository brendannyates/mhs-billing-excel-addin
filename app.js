import {EXTRA,filled,merge,closed,sla} from './core.js';
const $=s=>document.querySelector(s), esc=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const FORM='https://forms.cloud.microsoft/Pages/ResponsePage.aspx?id=8Sq2y8CkOEWdeGKInLNdYw1_rrxVxaFImLkzEWnt0GFUQUtJVFg2VzU5UEowM01JMUgzUUpKMTZPVi4u';
let rows=[],headers=[],refs=[],holidays=[],current='my',filter='',editing=null,busy=false;
const options={'Department Submitting Ticket:':2,'Patient Clinic:':3,'Urgency Level:':5,'Source of Inquiry:':6,'Task Type:':7,'Status:':10,'Was false verification of balance/charge given by Arietis?':14,'Patient Feedback/Service Recovery Flag:':15,'Source of Error:':16,'Outcome:':17};
function objects(v){return v.slice(1).filter(r=>filled(r[0])).map((r,index)=>({...Object.fromEntries(v[0].map((h,i)=>[h,r[i]??''])),_row:index}));}
async function run(fn){if(busy)return;busy=true;try{await fn();$('#message').textContent='Updated. All ticket edits are saved on Master.';}catch(e){$('#message').textContent=e.message;}finally{busy=false;}}
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
  if(!s.isNullObject){const v=s.getUsedRange();v.load('values');await ctx.sync();holidays=v.values.slice(1).map(x=>String(x[3]||'')).filter(x=>/^\d{4}-\d{2}-\d{2}$/.test(x));}
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
    const flags=sla(out.next,holidays);out.next['Confirmation Due']=flags.confirmation;out.next['Resolution Due']=flags.resolution;out.next['Escalation Flag']=flags.flags.join('; ');out.next['Sync Conflicts']=[...pending].join('; ');
    if(pos===undefined){ws.getRangeByIndexes(append,0,1,mh.length).values=[mh.map(h=>out.next[h]??'')];index.set(id,append++);}
    else for(let j=0;j<mh.length;j++)if(String(out.next[mh[j]]??'')!==String(mr.values[pos][j]??''))ws.getCell(pos,j).values=[[out.next[mh[j]]??'']];
    states.set(id,[id,JSON.stringify(r),JSON.stringify(overrides),JSON.stringify([...pending])]);
  }
  if(states.size)state.getRangeByIndexes(1,0,states.size,4).values=[...states.values()];
  ws.tables.load('items/name');await ctx.sync();if(ws.tables.items.length)ws.tables.items[0].resize(ws.getRangeByIndexes(0,0,append,mh.length));await ctx.sync();
});await refresh();}
function filtered(){if(current==='my')return rows.filter(r=>String(r['Owner Email']||r.Email).toLowerCase()===$('#email').value.trim().toLowerCase());if(current==='clinic')return rows.filter(r=>!filter||r['Patient Clinic:']===filter);if(current==='mrn')return rows.filter(r=>filter&&String(r['MRN:']).trim()===filter.trim());if(current==='action')return rows.filter(r=>sla(r,holidays).flags.length||filled(r['Sync Conflicts']));return rows;}
function render(){document.querySelectorAll('[data-view]').forEach(b=>b.classList.toggle('active',b.dataset.view===current));
  if(current==='form'){$('#view').innerHTML=`<h2>Submit a billing ticket</h2><p>Forms writes to Raw Data. Sync promotes the submission to Master.</p><p><a href="${FORM}" target="_blank" rel="noopener noreferrer">Open form in a new window</a></p><iframe title="Billing escalation submission form" src="${FORM}&embed=true" width="640" height="480" frameborder="0" marginwidth="0" marginheight="0" style="border:none;max-width:100%;max-height:100vh" allowfullscreen></iframe>`;return;}
  if(current==='settings'){$('#view').innerHTML='<h2>Automation settings</h2><p>Edit holiday dates in Settings column D as YYYY-MM-DD. Business hours default to 9 AM–5 PM Pacific; configure the scheduled flows to match.</p><p class="notice">My tickets is a view filter, not a permission boundary. Workbook permissions control access.</p><p>All edits use Master. Raw Data and its Forms table must remain intact.</p>';return;}
  if(current==='stats'){const clinics=[...new Set([...refs.map(r=>r[0]),...rows.map(r=>r['Patient Clinic:'])].filter(filled))];$('#view').innerHTML='<h2>Clinic dashboard</h2><div class="metrics"><div class="card">Total tickets<strong>'+rows.length+'</strong></div><div class="card">Open<strong>'+rows.filter(r=>!closed(r)).length+'</strong></div></div><div class="scroll"><table><thead><tr><th>Clinic</th><th>Total</th><th>Open</th><th>Receipt overdue</th><th>Resolution overdue</th><th>Closed</th></tr></thead><tbody>'+clinics.map(c=>{const a=rows.filter(r=>r['Patient Clinic:']===c);return '<tr>'+[c,a.length,a.filter(r=>!closed(r)).length,a.filter(r=>sla(r,holidays).flags.includes('Receipt SLA overdue')).length,a.filter(r=>sla(r,holidays).flags.includes('Resolution SLA overdue')).length,a.filter(closed).length].map(x=>'<td>'+esc(x)+'</td>').join('')+'</tr>';}).join('')+'</tbody></table></div>';return;}
  let top='<h2>'+({my:'My tickets',clinic:'Clinic tickets',mrn:'MRN lookup',action:'Needs action'}[current])+'</h2>';
  if(current==='clinic')top+='<label>Clinic<select id="filter"><option value="">All clinics</option>'+[...new Set(refs.map(r=>r[0]).filter(filled))].map(c=>`<option ${c===filter?'selected':''}>${esc(c)}</option>`).join('')+'</select></label>';
  if(current==='mrn')top+=`<label>Exact MRN<input id="filter" value="${esc(filter)}" inputmode="numeric"></label>`;
  const a=filtered();$('#view').innerHTML=top+(a.length?a.map(r=>{const x=sla(r,holidays);return `<article class="card"><strong>Ticket #${esc(r.Id)} · ${esc(r['Patient Clinic:'])}</strong><p class="meta">${esc(r['Status:'])} · ${esc(r['Urgency Level:'])}</p><p>MRN: ${esc(r['MRN:'])} · ${esc(r['Task Type:'])}</p><p>Receipt due: ${esc(x.confirmation||'Set vendor submission time')}<br>Resolution due: ${esc(x.resolution||'—')}</p><p class="flag">${esc([...x.flags,r['Sync Conflicts']?'Review sync conflicts: '+r['Sync Conflicts']:''].filter(Boolean).join(' · '))}</p><button data-id="${esc(r.Id)}">Edit ticket</button></article>`;}).join(''):'<p class="empty">No matching tickets.</p>');
  if($('#filter'))$('#filter').onchange=e=>{filter=e.target.value;render();};document.querySelectorAll('[data-id]').forEach(b=>b.onclick=()=>edit(b.dataset.id));
}
function edit(id){editing=rows.find(r=>String(r.Id)===id);const fields=headers.filter(h=>!['Id','Start time','Completion time','Confirmation Due','Resolution Due','Escalation Flag','Sync Conflicts','Column1','Date of first reply from Arietis:','Last Reply At','Closed At'].includes(h));$('#fields').innerHTML=fields.map(h=>{const v=editing[h]??'';const opts=options[h]===undefined?null:[...new Set(refs.map(r=>r[options[h]]).filter(x=>filled(x)&&x!=='[Free Text]'))];let control;if(opts&&h!=='Task Type:'){const all=[...new Set(['',v,...opts])];control=`<select name="${esc(h)}">${all.map(x=>`<option ${String(x)===String(v)?'selected':''}>${esc(x)}</option>`).join('')}</select>`;}else control=`<textarea name="${esc(h)}">${esc(v)}</textarea>`;return `<label>${esc(h)}${control}</label>`;}).join('');$('#editor').showModal();}
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
Office.onReady(info=>{if(info.host!==Office.HostType.Excel)return;$('#message').textContent='Ready. Initialize once to enable the Master workflow.';$('#setup').onclick=()=>run(initialize);$('#refresh').onclick=()=>run(refresh);$('#sync').onclick=()=>run(sync);$('#email').onchange=render;document.querySelectorAll('[data-view]').forEach(b=>b.onclick=()=>{current=b.dataset.view;filter='';render();});$('#cancel').onclick=()=>$('#editor').close();$('#editForm').onsubmit=e=>{e.preventDefault();run(save);};run(refresh);});
