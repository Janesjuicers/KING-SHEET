/** Refreshes the 24-profile matrix while preserving statuses by base bookmaker name. */
function refreshBookieHealth() {
  const settings=mbSheet_('Settings'); setupBookieHealthSettings_(settings);
  const sheet=mbSheet_('Bookie Health'), preserved={}, lastRow=sheet.getLastRow(), lastCol=sheet.getLastColumn();
  if(lastRow>1 && lastCol>1) {
    const twoHeaders=String(sheet.getRange(2,2).getDisplayValue()).toLowerCase()==='status', first=twoHeaders?3:2;
    if(lastRow>=first) {
      const names=sheet.getRange(first,1,lastRow-first+1,1).getDisplayValues().flat();
      const data=sheet.getRange(first,2,lastRow-first+1,lastCol-1).getValues();
      names.forEach((name,i)=>{const base=baseBookmakerName_(name);if(base&&!preserved[base.toLowerCase()])preserved[base.toLowerCase()]=twoHeaders?data[i].filter((_,j)=>j%2===0):data[i];});
    }
  }
  const names=bookieHealthBookmakers_(), profiles=MB.BOOKIE_HEALTH_PROFILES, rows=MB.BOOKIE_HEALTH_ROWS;
  mbEnsureSize_(sheet,rows+2,profiles*2+1);
  sheet.getRange(1,1,2,profiles*2+1).breakApart().clearContent();
  sheet.getRange(1,1,2,1).merge().setValue('Bookmaker');
  for(let p=1;p<=profiles;p++){const c=2+(p-1)*2;sheet.getRange(1,c,1,2).merge().setValue('Profile '+p);sheet.getRange(2,c,1,2).setValues([['Status','Account ID']]);}
  sheet.getRange(3,1,rows,profiles*2+1).clearContent().clearDataValidations();
  if(names.length){
    sheet.getRange(3,1,names.length,1).setValues(names.map(x=>[x]));
    const matrix=names.map(name=>{const saved=preserved[name.toLowerCase()]||[];return Array.from({length:profiles},(_,p)=>[saved[p]||'',saved[p]?bookieAccountId_(p+1,name):'']).flat();});
    sheet.getRange(3,2,names.length,profiles*2).setValues(matrix);
  }
  refreshBookieHealthValidation_(sheet,rows,profiles); protectBookieHealthIds_(sheet,rows,profiles);
  formatBookieHealth_(sheet,rows,profiles); refreshAccountIdList_(); refreshAccounts_(); refreshAccountValidations_();
  SpreadsheetApp.getActive().toast('Bookie Health refreshed.');
}

function baseBookmakerName_(name) {
  let value=String(name||'').trim();
  const match=value.match(/^(?:([1-9]|1\d|2[0-4]))?(BET365|TAB|SPORTSBET|UNIBET|BETNATION|NEXTBET)$/i);
  if(match){const canonical={BET365:'Bet365',TAB:'TAB',SPORTSBET:'Sportsbet',UNIBET:'Unibet',BETNATION:'BetNation',NEXTBET:'NextBet'};value=canonical[match[2].toUpperCase()];}
  return value;
}
function bookieAccountId_(profile,name){const clean=bookieAccountName_(name);return clean?String(profile)+clean:'';}
function bookieAccountName_(name){return baseBookmakerName_(name).replace(/[^\p{L}\p{N}]/gu,'');}

function refreshAccountIdList_() {
  const health=mbSheet_('Bookie Health'),settings=mbSheet_('Settings'),ids=[];
  const count=Math.min(MB.BOOKIE_HEALTH_ROWS,Math.max(0,health.getLastRow()-2));
  if(count)for(let c=3;c<=MB.BOOKIE_HEALTH_PROFILES*2+1;c+=2)health.getRange(3,c,count,1).getDisplayValues().flat().forEach(id=>{if(id)ids.push(id);});
  const unique=[...new Set(ids)].sort((a,b)=>(parseInt(a,10)||0)-(parseInt(b,10)||0)||a.localeCompare(b));
  settings.getRange(1,23,settings.getMaxRows(),1).clearContent();settings.getRange('W1').setValue('Generated Account IDs (do not edit)');
  if(unique.length)settings.getRange(2,23,unique.length,1).setValues(unique.map(x=>[x]));
  SpreadsheetApp.getActive().setNamedRange('Account_ID_List',settings.getRange(2,23,Math.max(1,unique.length),1));settings.hideColumns(23);return unique;
}

/** Adds generated accounts without changing existing account information. */
function refreshAccounts_(){
  const s=mbSheet_('Accounts'),ids=refreshAccountIdList_();mbEnsureSize_(s,Math.max(MB.ROWS,ids.length+1),1);
  if(!s.getRange('A1').getValue())s.getRange('A1').setValue('Account');
  const headers=s.getRange(1,1,1,Math.max(1,s.getLastColumn())).getDisplayValues()[0],col=Math.max(1,headers.indexOf('Account')+1);
  const existing=s.getRange(2,col,Math.max(1,s.getLastRow()-1),1).getDisplayValues().flat(),seen=new Set(existing);
  const add=ids.filter(x=>!seen.has(x));if(add.length)s.getRange(Math.max(2,s.getLastRow()+1),col,add.length,1).setValues(add.map(x=>[x]));
}
function bookieHealthBookmakers_(){
  const seen=new Set(),result=[];
  mbSheet_('Settings').getRange(2,22,MB.BOOKIE_HEALTH_ROWS,1).getDisplayValues().flat().forEach(raw=>{const n=baseBookmakerName_(raw),key=bookieAccountName_(n).toLowerCase();if(n&&key&&!seen.has(key)){seen.add(key);result.push(n);}});return result;
}
function activeBookieHealthStatuses_(){return mbSheet_('Settings').getRange('P2:T100').getValues().filter(r=>String(r[0]).trim()&&r[4]===true).sort((a,b)=>(Number(a[3])||9999)-(Number(b[3])||9999));}
function refreshBookieHealthValidation_(sheet,rowCount,profileCount){
  const settings=mbSheet_('Settings'),statuses=activeBookieHealthStatuses_();settings.getRange('U1:U100').clearContent();settings.getRange('U1').setValue('Active Bookie Health Statuses');
  if(statuses.length)settings.getRange(2,21,statuses.length,1).setValues(statuses.map(r=>[r[0]]));
  const rule=SpreadsheetApp.newDataValidation().requireValueInRange(settings.getRange(2,21,Math.max(1,statuses.length),1),true).setAllowInvalid(false).setHelpText('Select a status; another selection appends without duplicates.').build();
  for(let p=0;p<profileCount;p++)sheet.getRange(3,2+p*2,rowCount,1).setDataValidation(rule);
}
function protectBookieHealthIds_(sheet,rows,profiles){
  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).filter(p=>p.getDescription()==='Managed Account ID').forEach(p=>p.remove());
  for(let p=0;p<profiles;p++)sheet.getRange(3,3+p*2,rows,1).protect().setDescription('Managed Account ID').setWarningOnly(true);
}
function formatBookieHealth_(sheet,rowCount,profileCount){
  sheet.setFrozenRows(2);sheet.setFrozenColumns(1);sheet.setColumnWidth(1,180);for(let c=2;c<=profileCount*2+1;c++)sheet.setColumnWidth(c,c%2===0?92:125);
  sheet.getRange(1,1,2,profileCount*2+1).setBackground('#fff200').setFontColor('#000').setFontWeight('bold').setHorizontalAlignment('center');
  const matrix=sheet.getRange(3,2,rowCount,profileCount*2).setHorizontalAlignment('center');matrix.setBorder(true,true,true,true,true,true,'#333',SpreadsheetApp.BorderStyle.SOLID);
  const statusRanges=Array.from({length:profileCount},(_,p)=>sheet.getRange(3,2+p*2,rowCount,1)),rules=[];
  activeBookieHealthStatuses_().forEach(st=>rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains(st[0]).setBackground(st[1]||'#fff').setFontColor(st[2]||'#000').setRanges(statusRanges).build()));sheet.setConditionalFormatRules(rules);sheet.getDataRange().setFontFamily('Arial');
}

/** Multi-select fallback, scoped solely to Bookie Health Status columns. */
function onEdit(e){
  if(!e||!e.range||e.range.getSheet().getName()!=='Bookie Health')return;
  const r=e.range;if(r.getNumRows()!==1||r.getNumColumns()!==1||r.getRow()<3||r.getRow()>MB.BOOKIE_HEALTH_ROWS+2||r.getColumn()<2||r.getColumn()%2!==0)return;
  let value='';
  if(e.value!==undefined){const selected=String(e.value).trim(),allowed=new Set(activeBookieHealthStatuses_().map(x=>String(x[0])));if(!allowed.has(selected))return;const values=String(e.oldValue||'').split(MB.BOOKIE_HEALTH_SEPARATOR).map(x=>x.trim()).filter(x=>allowed.has(x));if(!values.includes(selected))values.push(selected);value=values.join(MB.BOOKIE_HEALTH_SEPARATOR);r.setValue(value);}
  r.offset(0,1).setValue(value?bookieAccountId_((r.getColumn()/2),r.getSheet().getRange(r.getRow(),1).getDisplayValue()):'');
  refreshAccountIdList_();refreshAccounts_();refreshAccountValidations_();
}
