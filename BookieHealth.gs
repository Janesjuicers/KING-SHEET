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
    const matrix=names.map(name=>{const saved=preserved[name.toLowerCase()]||[];return Array.from({length:profiles},(_,p)=>[saved[p]||'','']).flat();});
    sheet.getRange(3,2,names.length,profiles*2).setValues(matrix);
  }
  setBookieAccountIdFormulas_(sheet,rows,profiles);
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

/** Rebuilds the calculated account ledger from generated IDs. Risk equals pending withdrawals plus positive cash/bonus held in BAN or StatDec accounts. */
function refreshAccounts_(){
  const s=mbSheet_('Accounts'),ids=refreshAccountIdList_();mbEnsureSize_(s,Math.max(MB.ROWS,ids.length+4),12);
  const filter=s.getFilter();if(filter)filter.remove();s.getRange(1,1,s.getMaxRows(),12).breakApart().clearContent().clearDataValidations();
  s.getRange('A1:L1').merge().setValue('ACCOUNTS — one row per active generated Account ID');
  s.getRange('A2:H2').setValues([['Cash total','','Bonus total','','Pending withdrawals','','Risk exposure','']]);
  s.getRange(4,1,1,12).setValues([MB.ACCOUNT_HEADERS]);
  if(ids.length)s.getRange(5,1,ids.length,1).setFormulas(ids.map((_,i)=>[`=INDEX(Account_ID_List,${i+1})`]));
  const funds=SpreadsheetApp.getActive().getSheetByName('Funds Management');
  const fundSum=(names,account)=>{if(!funds)return '0';const hs=mbHeaders_(funds),a=hs.indexOf('Account')+1||hs.indexOf('Account ID')+1,c=names.map(n=>hs.indexOf(n)+1).find(Boolean);return a&&c?`SUMIF('Funds Management'!${mbColA1_(a)}:${mbColA1_(a)},${account},'Funds Management'!${mbColA1_(c)}:${mbColA1_(c)})`:'0';};
  for(let i=0;i<ids.length;i++){
    const r=i+5,a=`A${r}`,tag=`B${r}`;
    s.getRange(r,2).setFormula(`=IF(${a}="","",IFNA(XLOOKUP(${a},FLATTEN('Bookie Health'!C3:AW252),FLATTEN('Bookie Health'!B3:AV252)),""))`);
    const dep=fundSum(['Deposit','Deposits','Amount Deposited'],a),pending=fundSum(['Pending Withdrawals','Pending Withdrawal'],a),approved=fundSum(['Approved Withdrawals','Approved Withdrawal'],a),bonus=fundSum(['Bonus Received','Ongoing Bonus Received'],a);
    s.getRange(r,3).setFormula(`=IF(${a}="","",${dep}-${pending}-${approved}+SUMIF(Promos!B:B,${a},Promos!I:I)+SUMIF(BTO!B:B,${a},BTO!J:J)+SUMIF(NonPromos!B:B,${a},NonPromos!K:K))`);
    s.getRange(r,4).setFormula(`=IF(${a}="","",${bonus}+SUMIF(Promos!B:B,${a},Promos!J:J)-SUMIF(BTO!B:B,${a},BTO!E:E))`);
    s.getRange(r,5).setFormula(`=IF(${a}="","",${dep})`);s.getRange(r,6).setFormula(`=IF(${a}="","",${pending})`);s.getRange(r,7).setFormula(`=IF(${a}="","",${approved})`);
    s.getRange(r,8).setFormula(`=IF(${a}="","",C${r}+F${r}+G${r}-E${r})`);
    s.getRange(r,9).setFormula(`=IF(${a}="","",F${r}+IF(REGEXMATCH(${tag},"BAN|StatDec"),MAX(0,C${r})+MAX(0,D${r}),0))`);
    s.getRange(r,10).setFormula(`=IF(${a}="","",COUNTIFS(Promos!B:B,${a},Promos!J:J,">0"))`);s.getRange(r,11).setFormula(`=IF(${a}="","",${bonus})`);
    s.getRange(r,12).setFormula(`=IF(${a}="","",IF(REGEXMATCH(${tag},"BAN"),"banned",IF(REGEXMATCH(${tag},"StatDec"),"restricted/review",IF(ABS(C${r})+ABS(D${r})+E${r}+F${r}+G${r}>0,"active","active-empty"))))`);
  }
  s.getRange('B2').setFormula('=SUM(C5:C)');s.getRange('D2').setFormula('=SUM(D5:D)');s.getRange('F2').setFormula('=SUM(F5:F)');s.getRange('H2').setFormula('=SUM(I5:I)');
  s.getRange(4,1,s.getMaxRows()-3,12).createFilter();s.setFrozenRows(4);
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
function setBookieAccountIdFormulas_(sheet,rows,profiles){
  for(let p=1;p<=profiles;p++){
    const status=mbColA1_(p*2);
    sheet.getRange(3,p*2+1,rows,1).clearContent().clearDataValidations();
    sheet.getRange(3,p*2+1).setFormula(`=ARRAYFORMULA(IF(${status}3:${status}="","",${p}&REGEXREPLACE(A3:A,"[^[:alnum:]]","")))`);
  }
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
  SpreadsheetApp.flush();
  refreshAccountIdList_();refreshAccounts_();refreshAccountValidations_();
}
