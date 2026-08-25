/** Refreshes the two-column-per-profile matrix, preserving statuses by bookmaker/profile. */
function refreshBookieHealth() {
  const settings=mbSheet_('Settings');
  setupBookieHealthSettings_(settings);
  const sheet=mbSheet_('Bookie Health');
  const oldLastRow=sheet.getLastRow(), oldLastCol=sheet.getLastColumn();
  const preserved={};
  if (oldLastRow>1 && oldLastCol>1) {
    const twoHeaders=String(sheet.getRange(2,2).getDisplayValue()).toLowerCase()==='status';
    const firstDataRow=twoHeaders?3:2;
    const oldNames=sheet.getRange(firstDataRow,1,oldLastRow-firstDataRow+1,1).getDisplayValues();
    const oldData=sheet.getRange(firstDataRow,2,oldLastRow-firstDataRow+1,oldLastCol-1).getValues();
    oldNames.forEach((r,i)=>{const name=String(r[0]).trim();if(name && !preserved[name]) preserved[name]=twoHeaders?oldData[i].filter((_,j)=>j%2===0):oldData[i];});
  }
  const names=bookieHealthBookmakers_();
  const cleanCounts={};
  const idNames=names.reduce((m,n)=>{const clean=bookieAccountName_(n), count=(cleanCounts[clean]||0)+1;cleanCounts[clean]=count;m[n]=clean+(count>1?'_'+count:'');return m;},{});
  const existingProfiles=Math.max(0,Math.floor((oldLastCol-1)/(String(sheet.getRange(2,2).getDisplayValue()).toLowerCase()==='status'?2:1)));
  const profileCount=Math.max(MB.BOOKIE_HEALTH_PROFILES,existingProfiles);
  mbEnsureSize_(sheet,Math.max(100,names.length+2),profileCount*2+1);
  sheet.getRange(1,1,2,profileCount*2+1).breakApart().clearContent();
  sheet.getRange(1,1,2,1).merge().setValue('Bookmaker');
  for(let p=1;p<=profileCount;p++) { const c=2+(p-1)*2; sheet.getRange(1,c,1,2).merge().setValue('Profile '+p); sheet.getRange(2,c,1,2).setValues([['Status','Account ID']]); }
  if (oldLastRow>1) sheet.getRange(3,1,Math.max(1,oldLastRow-2),profileCount*2+1).clearContent().clearDataValidations();
  if (names.length) {
    sheet.getRange(3,1,names.length,1).setValues(names.map(n=>[n]));
    sheet.getRange(3,2,names.length,profileCount*2).setValues(names.map(n=>{
      const row=(preserved[n]||[]).slice(0,profileCount);
      while(row.length<profileCount) row.push('');
      return row.reduce((out,status,p)=>out.concat([status,String(p+1)+idNames[n]]),[]);
    }));
  }
  refreshBookieHealthValidation_(sheet,names.length,profileCount);
  formatBookieHealth_(sheet,names.length,profileCount);
  refreshAccountIdList_();
  refreshAccounts_();
  refreshAccountValidations_();
  SpreadsheetApp.getActive().toast('Bookie Health refreshed.');
}

function bookieAccountId_(profile,name) {
  const clean=bookieAccountName_(name);
  return clean ? String(profile)+clean : '';
}

function bookieAccountName_(name) { return String(name||'').replace(/\s+/g,'').replace(/[^\p{L}\p{N}_-]/gu,''); }

function refreshAccountIdList_() {
  const health=mbSheet_('Bookie Health'), settings=mbSheet_('Settings');
  const ids=[];
  if(health.getLastRow()>2) for(let c=3;c<=health.getLastColumn();c+=2) health.getRange(3,c,health.getLastRow()-2,1).getDisplayValues().forEach(r=>{if(r[0])ids.push(r[0]);});
  const unique=[...new Set(ids)].sort((a,b)=>{const pa=parseInt(a,10)||0,pb=parseInt(b,10)||0;return pa-pb||a.replace(/^\d+/,'').localeCompare(b.replace(/^\d+/,''),undefined,{sensitivity:'base'});});
  settings.getRange('W:W').clearContent(); settings.getRange('W1').setValue('Generated Account IDs (do not edit)');
  if(unique.length) settings.getRange(2,23,unique.length,1).setValues(unique.map(x=>[x]));
  const ss=SpreadsheetApp.getActive(), old=ss.getRangeByName('Account_ID_List');
  ss.setNamedRange('Account_ID_List',settings.getRange(2,23,Math.max(1,unique.length),1));
  settings.hideColumns(23);
  return unique;
}

/** Adds missing accounts without changing or reordering existing account data. */
function refreshAccounts_() {
  const s=mbSheet_('Accounts'), ids=refreshAccountIdList_(); mbEnsureSize_(s,Math.max(MB.ROWS,ids.length+1),1);
  if(!s.getRange('A1').getValue()) s.getRange('A1').setValue('Account');
  let accountCol=Math.max(1,s.getRange(1,1,1,Math.max(1,s.getLastColumn())).getDisplayValues()[0].findIndex(x=>String(x).trim()==='Account')+1);
  const existing=s.getRange(2,accountCol,Math.max(1,s.getLastRow()-1),1).getDisplayValues().flat().filter(Boolean), seen=new Set(existing);
  const add=ids.filter(x=>!seen.has(x)); if(add.length)s.getRange(Math.max(2,s.getLastRow()+1),accountCol,add.length,1).setValues(add.map(x=>[x]));
}

function bookieHealthBookmakers_() {
  const values=mbSheet_('Settings').getRange('V2:V').getDisplayValues().flat().map(v=>v.trim()).filter(Boolean);
  return [...new Set(values)];
}

function activeBookieHealthStatuses_() {
  return mbSheet_('Settings').getRange('P2:T100').getValues()
    .filter(r=>String(r[0]).trim() && r[4]===true)
    .sort((a,b)=>(Number(a[3])||9999)-(Number(b[3])||9999));
}

function refreshBookieHealthValidation_(sheet,rowCount,profileCount) {
  const settings=mbSheet_('Settings'), statuses=activeBookieHealthStatuses_();
  settings.getRange('U1:U100').clearContent();
  settings.getRange('U1').setValue('Active Bookie Health Statuses');
  if(statuses.length) settings.getRange(2,21,statuses.length,1).setValues(statuses.map(r=>[r[0]]));
  if(!rowCount) return;
  const source=settings.getRange(2,21,Math.max(1,statuses.length),1);
  const rule=SpreadsheetApp.newDataValidation().requireValueInRange(source,true).setAllowInvalid(false).setHelpText('Choose a status; choose again to add another status.').build();
  for(let p=0;p<profileCount;p++) sheet.getRange(3,2+p*2,rowCount,1).setDataValidation(rule);
}

function formatBookieHealth_(sheet,rowCount,profileCount) {
  sheet.setFrozenRows(2); sheet.setFrozenColumns(1);
  sheet.setRowHeights(1,2,28); if(rowCount) sheet.setRowHeights(3,rowCount,24);
  sheet.setColumnWidth(1,180); for(let c=2;c<=profileCount*2+1;c++) sheet.setColumnWidth(c,c%2===0?92:125);
  sheet.getRange(1,1,2,profileCount*2+1).setBackground('#fff200').setFontColor('#000000').setFontWeight('bold').setHorizontalAlignment('center');
  if(rowCount) {
    const labels=sheet.getRange(3,1,rowCount,1).setFontWeight('bold');
    labels.setBackgrounds(Array.from({length:rowCount},(_,i)=>[i%2?'#d9ead3':'#cfe2f3']));
    const matrix=sheet.getRange(3,2,rowCount,profileCount*2).setHorizontalAlignment('center').setVerticalAlignment('middle');
    matrix.setBorder(true,true,true,true,true,true,'#333333',SpreadsheetApp.BorderStyle.SOLID);
    const rules=sheet.getConditionalFormatRules().filter(r=>!r.getRanges().some(x=>x.getRow()>=2 && x.getColumn()>=2));
    const statusRanges=Array.from({length:profileCount},(_,p)=>sheet.getRange(3,2+p*2,rowCount,1));
    activeBookieHealthStatuses_().forEach(st=>rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains(st[0]).setBackground(st[1]||'#ffffff').setFontColor(st[2]||'#000000').setRanges(statusRanges).build()));
    sheet.setConditionalFormatRules(rules);
  }
  sheet.getDataRange().setFontFamily('Arial');
}

/** Apps Script has no reliable API for enabling native multi-select chips, so append selections safely. */
function onEdit(e) {
  if(!e || !e.range || e.range.getSheet().getName()!=='Bookie Health') return;
  const r=e.range;
  if(r.getNumRows()!==1 || r.getNumColumns()!==1 || r.getRow()<3 || r.getColumn()<2 || r.getColumn()%2!==0 || e.value===undefined) return;
  const selected=String(e.value).trim(), old=String(e.oldValue||'').trim();
  if(!selected) return;
  const allowed=new Set(activeBookieHealthStatuses_().map(x=>String(x[0])));
  if(!allowed.has(selected)) return;
  const values=old ? old.split(MB.BOOKIE_HEALTH_SEPARATOR).map(x=>x.trim()).filter(Boolean) : [];
  if(!values.includes(selected)) values.push(selected);
  r.setValue(values.join(MB.BOOKIE_HEALTH_SEPARATOR));
}
