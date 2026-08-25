/** Refreshes only the Bookie Health matrix, preserving selections by bookmaker name. */
function refreshBookieHealth() {
  const settings=mbSheet_('Settings');
  setupBookieHealthSettings_(settings);
  const sheet=mbSheet_('Bookie Health');
  const oldLastRow=sheet.getLastRow(), oldLastCol=sheet.getLastColumn();
  const preserved={};
  if (oldLastRow>1 && oldLastCol>1) {
    const oldNames=sheet.getRange(2,1,oldLastRow-1,1).getDisplayValues();
    const oldData=sheet.getRange(2,2,oldLastRow-1,oldLastCol-1).getValues();
    oldNames.forEach((r,i)=>{const name=String(r[0]).trim();if(name && !preserved[name]) preserved[name]=oldData[i];});
  }
  const names=bookieHealthBookmakers_();
  const existingProfiles=Math.max(0,oldLastCol-1);
  const profileCount=Math.max(MB.BOOKIE_HEALTH_PROFILES,existingProfiles);
  mbEnsureSize_(sheet,Math.max(100,names.length+1),profileCount+1);
  sheet.getRange(1,1,1,profileCount+1).setValues([['Bookmaker'].concat(Array.from({length:profileCount},(_,i)=>i+1))]);
  if (oldLastRow>1) sheet.getRange(2,1,oldLastRow-1,profileCount+1).clearContent();
  if (names.length) {
    sheet.getRange(2,1,names.length,1).setValues(names.map(n=>[n]));
    sheet.getRange(2,2,names.length,profileCount).setValues(names.map(n=>{
      const row=(preserved[n]||[]).slice(0,profileCount);
      while(row.length<profileCount) row.push('');
      return row;
    }));
  }
  refreshBookieHealthValidation_(sheet,names.length,profileCount);
  formatBookieHealth_(sheet,names.length,profileCount);
  SpreadsheetApp.getActive().toast('Bookie Health refreshed.');
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
  sheet.getRange(2,2,rowCount,profileCount).setDataValidation(rule);
}

function formatBookieHealth_(sheet,rowCount,profileCount) {
  sheet.setFrozenRows(1); sheet.setFrozenColumns(1);
  sheet.setRowHeight(1,30); if(rowCount) sheet.setRowHeights(2,rowCount,24);
  sheet.setColumnWidth(1,180); for(let c=2;c<=profileCount+1;c++) sheet.setColumnWidth(c,92);
  sheet.getRange(1,1,1,profileCount+1).setBackground('#fff200').setFontColor('#000000').setFontWeight('bold').setHorizontalAlignment('center');
  if(rowCount) {
    const labels=sheet.getRange(2,1,rowCount,1).setFontWeight('bold');
    labels.setBackgrounds(Array.from({length:rowCount},(_,i)=>[i%2?'#d9ead3':'#cfe2f3']));
    const matrix=sheet.getRange(2,2,rowCount,profileCount).setHorizontalAlignment('center').setVerticalAlignment('middle');
    matrix.setBorder(true,true,true,true,true,true,'#333333',SpreadsheetApp.BorderStyle.SOLID);
    const rules=sheet.getConditionalFormatRules().filter(r=>!r.getRanges().some(x=>x.getRow()>=2 && x.getColumn()>=2));
    activeBookieHealthStatuses_().forEach(st=>rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(st[0]).setBackground(st[1]||'#ffffff').setFontColor(st[2]||'#000000').setRanges([matrix]).build()));
    sheet.setConditionalFormatRules(rules);
  }
  sheet.getDataRange().setFontFamily('Arial');
}

/** Apps Script has no reliable API for enabling native multi-select chips, so append selections safely. */
function onEdit(e) {
  if(!e || !e.range || e.range.getSheet().getName()!=='Bookie Health') return;
  const r=e.range;
  if(r.getNumRows()!==1 || r.getNumColumns()!==1 || r.getRow()<2 || r.getColumn()<2 || e.value===undefined) return;
  const selected=String(e.value).trim(), old=String(e.oldValue||'').trim();
  if(!selected) return;
  const allowed=new Set(activeBookieHealthStatuses_().map(x=>String(x[0])));
  if(!allowed.has(selected)) return;
  const values=old ? old.split(MB.BOOKIE_HEALTH_SEPARATOR).map(x=>x.trim()).filter(Boolean) : [];
  if(!values.includes(selected)) values.push(selected);
  r.setValue(values.join(MB.BOOKIE_HEALTH_SEPARATOR));
}
