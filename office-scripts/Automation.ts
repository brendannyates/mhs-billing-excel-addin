/** Run from Power Automate. Serialize every workbook writer through one worker flow. */
type Cell = string | number | boolean;
type Ticket = {[key: string]: Cell};
type Reply = {conversationId: string; sender: string; receivedAt: string; messageId: string};
type State = {baseline: Ticket; overrides: string[]; conflicts: string[]};
function present(v: Cell | undefined): boolean { return v !== undefined && String(v).trim() !== ''; }
function isClosed(t: Ticket): boolean { return /^(Closed|Resolved)\b/.test(String(t['Status:']||'')); }
function day(v: Cell): string {
  if(typeof v==='number')return new Date(Math.round((v-25569)*86400000)).toISOString().slice(0,10);
  if(/^\d{4}-\d{2}-\d{2}$/.test(String(v)))return String(v);
  const d=new Date(String(v));if(isNaN(d.getTime()))return '';
  // Power Automate passes Pacific local date explicitly for SLA comparisons.
  return new Intl.DateTimeFormat('en-CA',{timeZone:'America/Los_Angeles',year:'numeric',month:'2-digit',day:'2-digit'}).format(d);
}
function due(start: Cell, count: number, holidays: string[]): string {
  const s=day(start);if(!s)return '';const d=new Date(s+'T12:00:00Z');let n=0;
  while(n<count){d.setUTCDate(d.getUTCDate()+1);if(d.getUTCDay()!==0&&d.getUTCDay()!==6&&!holidays.includes(d.toISOString().slice(0,10)))n++;}
  return d.toISOString().slice(0,10);
}
function businessAge(start: Cell, today: string, holidays: string[]): number {
  const key=day(start);if(!key||key>today)return 0;const d=new Date(key+'T12:00:00Z');let n=0;
  while(d.toISOString().slice(0,10)<today){d.setUTCDate(d.getUTCDate()+1);if(d.getUTCDay()!==0&&d.getUTCDay()!==6&&!holidays.includes(d.toISOString().slice(0,10)))n++;}return n;
}
function object(headers: string[], row: Cell[]): Ticket {const o: Ticket={};headers.forEach((h,i)=>o[h]=row[i]??'');return o;}
function epoch(v: Cell): number {return typeof v==='number'?(v-25569)*86400000:Date.parse(String(v));}
function main(workbook: ExcelScript.Workbook, operation: string='sync', localToday: string='', repliesJson: string='[]', mode: string='open'): string {
  if(!/^\d{4}-\d{2}-\d{2}$/.test(localToday))throw Error('Pass localToday as Pacific YYYY-MM-DD.');
  const master=workbook.getWorksheet('Master'),raw=workbook.getWorksheet('Raw Data'),stateSheet=workbook.getWorksheet('_SyncState'),settings=workbook.getWorksheet('Settings');
  if(!master||!raw||!stateSheet||!settings)throw Error('Initialize with Excel add-in first.');
  const values=master.getUsedRange().getValues(),headers=values[0].map(String),tickets=values.slice(1).filter(r=>present(r[0])).map(r=>object(headers,r));
  const duplicate=new Set<string>();tickets.forEach(t=>{if(duplicate.has(String(t.Id)))throw Error('Duplicate Master Id');duplicate.add(String(t.Id));});
  const sv=stateSheet.getUsedRange().getValues(),states: {[id: string]: State}={};sv.slice(1).forEach(r=>{if(present(r[0]))states[String(r[0])]={baseline:JSON.parse(String(r[1]||'{}')),overrides:JSON.parse(String(r[2]||'[]')),conflicts:JSON.parse(String(r[3]||'[]'))};});
  const holidays=settings.getUsedRange().getValues().slice(1).map(r=>String(r[3]||'')).filter(v=>/^\d{4}-\d{2}-\d{2}$/.test(v));
  const unmatched: string[]=[];
  if(operation==='sync'){
    const rv=raw.getUsedRange().getValues(),rh=rv[0].map(String),seen=new Set<string>();
    rv.slice(1).forEach(r=>{if(present(r[0])){const id=String(r[0]);if(seen.has(id))throw Error('Duplicate Raw Data Id');seen.add(id);}});
    rv.slice(1).forEach(r=>{
      if(!present(r[0]))return;const source=object(rh,r),id=String(r[0]);let target=tickets.find(t=>String(t.Id)===id);
      if(!target){target={};headers.forEach(h=>target[h]='');tickets.push(target);}
      const st=states[id]||{baseline:{},overrides:[],conflicts:[]};
      rh.forEach(h=>{
        if(!headers.includes(h)||!present(source[h])||st.overrides.includes(h))return;
        if(h==='Date of first reply from Arietis:' && present(target[h]))return;
        if(h==='Status:' && isClosed(target))return;
        if(h==='Status:' && /^(Closed|Resolved)\b/.test(String(source[h]))&&!isClosed(target)){st.conflicts.push('Status: (close after inbox reconciliation)');return;}
        if(!present(target[h]))target[h]=source[h];
        else if(String(source[h])!==String(st.baseline[h]??'')&&String(source[h])!==String(target[h])){
          if(st.baseline[h]!==undefined&&String(target[h])===String(st.baseline[h]))target[h]=source[h];else st.conflicts.push(h);
        }
      });
      target['Owner Email']=target['Owner Email']||target.Email||'';st.baseline=source;st.conflicts=[...new Set(st.conflicts)];target['Sync Conflicts']=st.conflicts.join('; ');states[id]=st;
    });
  }else if(operation==='reply'){
    const replies: Reply[]=JSON.parse(repliesJson);
    replies.forEach(r=>{
      if(r.sender.toLowerCase()!=='patientbilling@arietishealth.com')return;
      if(!r.conversationId||!Number.isFinite(Date.parse(r.receivedAt)))throw Error('Invalid reply payload');
      const matches=tickets.filter(t=>String(t['Vendor Conversation ID'])===r.conversationId);
      if(matches.length!==1){unmatched.push(r.messageId);return;}
      const t=matches[0],received=Date.parse(r.receivedAt),cutoff=present(t['Closed At'])?epoch(t['Closed At']):Infinity;
      if(received>cutoff)return;
      const first=t['Date of first reply from Arietis:'],last=t['Last Reply At'];
      if(!present(first)||received<epoch(first))t['Date of first reply from Arietis:']=r.receivedAt;
      if(!present(last)||received>epoch(last))t['Last Reply At']=r.receivedAt;
      if(!isClosed(t)&&t['Status:']==='MHS - Submitted to Arietis')t['Status:']='Arietis - Confirmed Receipt';
      const st=states[String(t.Id)]||{baseline:{},overrides:[],conflicts:[]};st.overrides=[...new Set([...st.overrides,'Date of first reply from Arietis:','Status:'])];states[String(t.Id)]=st;
    });
  }else if(operation!=='digest')throw Error('Unknown operation');
  tickets.forEach(t=>{
    t['Confirmation Due']=due(t['Submitted to Vendor At'],2,holidays);t['Resolution Due']=due(t['Submitted to Vendor At'],5,holidays);
    const flags: string[]=[];if(!isClosed(t)){
      if(!present(t['Submitted to Vendor At']))flags.push('Vendor submission time needed');
      if(present(t['Confirmation Due'])&&!present(t['Date of first reply from Arietis:'])&&localToday>String(t['Confirmation Due']))flags.push('Receipt SLA overdue');
      if(present(t['Resolution Due'])&&localToday>String(t['Resolution Due']))flags.push('Resolution SLA overdue');
      if(t['Urgency Level:']==='Critical')flags.push('Critical priority');
    }t['Escalation Flag']=flags.join('; ');
    const reminderFlags: string[]=[],entered=t['Completion time']||t['Start time']||'';
    const dates=[t['Follow Up At'],t['Date of Patient Outreach (if applicable):'],t['Last Reply At'],t['Date of first reply from Arietis:']].filter(present).map(day).filter(Boolean).sort();
    if(!isClosed(t)){
      if(!present(t['Date of first reply from Arietis:'])&&businessAge(entered,localToday,holidays)>3)reminderFlags.push('Needs follow-up beyond 3 business days');
      if(businessAge(dates[dates.length-1]||entered,localToday,holidays)>5)reminderFlags.push('No follow-up beyond 5 business days');
      if(businessAge(t['Submitted to Vendor At']||entered,localToday,holidays)>6)reminderFlags.push('Resolution overdue beyond 6 business days');
    }t['Reminder Flags']=reminderFlags.join('; ');
  });
  // Single serialized worker is mandatory: no independent flows may race this write.
  if(operation!=='digest'){
    if(tickets.length)master.getRangeByIndexes(1,0,tickets.length,headers.length).setValues(tickets.map(t=>headers.map(h=>t[h]??'')));
    const tables=master.getTables();if(tables.length)tables[0].resize(master.getRangeByIndexes(0,0,tickets.length+1,headers.length));
    const stateRows=Object.keys(states).map(id=>[id,JSON.stringify(states[id].baseline),JSON.stringify(states[id].overrides),JSON.stringify(states[id].conflicts)]);
    if(stateRows.length)stateSheet.getRangeByIndexes(1,0,stateRows.length,4).setValues(stateRows);
  }
  const groups: {[email: string]: Ticket[]}={};
  tickets.forEach(t=>{
    const recipient=String(t['Owner Email']||t.Email||'').trim().toLowerCase();
    if(!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(recipient))return;
    if(mode!=='weekly'&&isClosed(t)&&day(t['Closed At'])!==localToday)return;
    if(mode==='weekly'&&isClosed(t)){const since=new Date(localToday+'T00:00:00Z');since.setUTCDate(since.getUTCDate()-7);if(day(t['Closed At'])<since.toISOString().slice(0,10))return;}
    if(!groups[recipient])groups[recipient]=[];
    // Email summary deliberately omits MRN, patient, notes, balances, attachments.
    groups[recipient].push({'Id':t.Id,'Patient Clinic:':t['Patient Clinic:'],'Status:':t['Status:'],'Confirmation Due':t['Confirmation Due'],'Resolution Due':t['Resolution Due'],'Escalation Flag':t['Escalation Flag'],'Sync Conflicts':t['Sync Conflicts'],'Reminder Flags':t['Reminder Flags']});
  });
  return JSON.stringify({processed:tickets.length,unmatched,digests:Object.keys(groups).map(email=>({email,tickets:groups[email]}))});
}
