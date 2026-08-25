function buildWorkbook() {
  MB.SHEETS.forEach(n => mbSheet_(n));
  setupSettings_();
  migrateEntrySheet_('Promos',MB.PROMO_HEADERS);
  migrateBtoSheet_();
  migrateEntrySheet_('NonPromos',MB.NON_HEADERS);
  refreshBookieHealth();
  seedPromoExamples_();
  refreshFormulas();
  refreshDropdowns();
  refreshDashboard();
  formatWorkbook_();
  runAudit();
  SpreadsheetApp.getActive().toast('Matched-betting workbook is ready.');
}

/** Maps existing BTO values by header so removing Entry ID/Month cannot shift entries. */
function migrateBtoSheet_() {
  const s=mbSheet_('BTO'), wanted=MB.BTO_HEADERS, lastCol=s.getLastColumn(), lastRow=s.getLastRow();
  if(!lastRow || !s.getRange(1,1).getValue()) { setupEntrySheet_(s,wanted); trimBtoColumns_(s); return; }
  const oldHeaders=s.getRange(1,1,1,lastCol).getDisplayValues()[0].map(x=>String(x).trim());
  if(oldHeaders.slice(0,wanted.length).join('\u0001')===wanted.join('\u0001') && !oldHeaders.includes('Entry ID') && !oldHeaders.includes('Month')) { setupEntrySheet_(s,wanted); trimBtoColumns_(s); return; }
  const rows=lastRow>1?s.getRange(2,1,lastRow-1,lastCol).getValues():[];
  const mapped=rows.map(row=>wanted.map(h=>{const i=oldHeaders.indexOf(h);return i<0?'':row[i];}));
  const filter=s.getFilter(); if(filter)filter.remove();
  s.getDataRange().clearContent(); mbEnsureSize_(s,MB.ROWS,wanted.length); s.getRange(1,1,1,wanted.length).setValues([wanted]);
  if(mapped.length)s.getRange(2,1,mapped.length,wanted.length).setValues(mapped);
  s.getRange(1,1,s.getMaxRows(),wanted.length).createFilter(); s.setFrozenRows(1);
  trimBtoColumns_(s);
}

/** Rebuilds a changed entry layout by header name, without moving one field into another. */
function migrateEntrySheet_(name,wanted) {
  const s=mbSheet_(name), lastRow=s.getLastRow(), lastCol=s.getLastColumn();
  if(!lastRow || !s.getRange(1,1).getValue()) { setupEntrySheet_(s,wanted); trimEntryColumns_(s,wanted.length); return; }
  const old=s.getRange(1,1,1,lastCol).getDisplayValues()[0].map(String);
  if(old.slice(0,wanted.length).join('\u0001')===wanted.join('\u0001') && old.length===wanted.length) { setupEntrySheet_(s,wanted); return; }
  const values=lastRow>1?s.getRange(2,1,lastRow-1,lastCol).getValues():[];
  const mapped=values.map(row=>wanted.map(h=>{const i=old.indexOf(h);return i<0?'':row[i];}));
  const filter=s.getFilter(); if(filter)filter.remove();
  s.getDataRange().clearContent(); mbEnsureSize_(s,MB.ROWS,wanted.length);
  s.getRange(1,1,1,wanted.length).setValues([wanted]);
  if(mapped.length)s.getRange(2,1,mapped.length,wanted.length).setValues(mapped);
  trimEntryColumns_(s,wanted.length); s.getRange(1,1,s.getMaxRows(),wanted.length).createFilter(); s.setFrozenRows(1);
}

function trimEntryColumns_(s,count) {
  if(s.getMaxColumns()>count)s.deleteColumns(count+1,s.getMaxColumns()-count);
}

function trimBtoColumns_(s) {
  if(s.getMaxColumns()<=MB.BTO_HEADERS.length)return;
  const filter=s.getFilter(); if(filter)filter.remove();
  s.deleteColumns(MB.BTO_HEADERS.length+1,s.getMaxColumns()-MB.BTO_HEADERS.length);
  s.getRange(1,1,s.getMaxRows(),MB.BTO_HEADERS.length).createFilter();
}

function seedPromoExamples_() {
  const s=mbSheet_('Promos');
  if (s.getRange('A2:A').getValues().some(r=>r[0]!=='')) return;
  const today=new Date();
  s.getRange('A2:G4').setValues([
    [today,'','SGM3 Leg','Starter example: edit or delete',100,2.10,'W'],
    [today,'','Free Hit Single','Starter example: edit or delete',50,4.00,'B'],
    [today,'','SRM','Starter example: edit or delete',25,3.20,'L']]);
}

function setupSettings_() {
  const s = mbSheet_('Settings');
  mbEnsureSize_(s, 251, 23);
  const headings = [['Promo Type','EV %','','EV Tier','Minimum %','Maximum %','','Legacy Accounts (unused)','BTO Methods','Non-Promo Types','Results','Account Statuses','Transaction Types','Banks']];
  s.getRange(1,1,1,14).setValues(headings);
  if (!s.getRange('A2').getValue()) s.getRange(2,1,MB.PROMOS.length,2).setValues(MB.PROMOS);
  if (!s.getRange('D2').getValue()) s.getRange('D2:F7').setValues([['A+',.31,1],['A',.20,.3099],['B',.15,.1999],['C',.10,.1499],['D',.05,.0999],['E',0,.0499]]);
  if (!s.getRange('I2').getValue()) s.getRange(2,9,MB.BTO_METHODS.length,1).setValues(MB.BTO_METHODS.map(x=>[x]));
  if (!s.getRange('J2').getValue()) s.getRange(2,10,MB.NON_TYPES.length,1).setValues(MB.NON_TYPES.map(x=>[x]));
  if (!s.getRange('K2').getValue()) s.getRange(2,11,MB.RESULTS.length,1).setValues(MB.RESULTS.map(x=>[x]));
  if (!s.getRange('L2').getValue()) s.getRange('L2:L4').setValues([['Open'],['Restricted'],['Closed']]);
  if (!s.getRange('M2').getValue()) s.getRange('M2:M4').setValues([['Deposit'],['Withdrawal'],['Adjustment']]);
  setupBookieHealthSettings_(s);
  s.getRange('D10').setValue('Default BTO rate when BSP is missing');
  if (s.getRange('E10').isBlank()) s.getRange('E10').setValue(.75);
  s.getRange('A:B').setNumberFormat('0.00%'); s.getRange('A:A').setNumberFormat('@');
  s.getRange('B2:B100').setNumberFormat('0.00%'); s.getRange('E2:F10').setNumberFormat('0.00%');
}

function setupBookieHealthSettings_(s) {
  s.getRange('P1:T1').setValues([['Bookie Health Statuses','Background colour','Text colour','Sort order','Active']]);
  if (s.getRange('P2:P').getValues().every(r=>r[0]==='')) {
    s.getRange(2,16,MB.BOOKIE_HEALTH_STATUSES.length,5).setValues(MB.BOOKIE_HEALTH_STATUSES);
  }
  s.getRange('T2:T100').insertCheckboxes();
  s.getRange('V1').setValue('Bookmakers');
  if (s.getRange('V2:V').getValues().every(r=>r[0]==='')) {
    const candidates=[];
    s.getRange('H2:H100').getDisplayValues().forEach(r=>candidates.push(r[0]));
    ['Promos','BTO','NonPromos'].forEach(n=>{const entry=mbSheet_(n), headers=entry.getRange(1,1,1,Math.max(1,entry.getLastColumn())).getDisplayValues()[0];const c=headers.indexOf('Account')+1;if(c)entry.getRange(2,c,MB.ROWS-1,1).getDisplayValues().forEach(r=>candidates.push(r[0]));});
    let names=[...new Set(candidates.map(baseBookmakerName_).filter(Boolean))];
    if(!names.length)names=['Bet365','TAB','Sportsbet','Unibet','BetNation','NextBet'];
    if (names.length) s.getRange(2,22,names.length,1).setValues(names.map(v=>[v]));
  }
  const bookmakerRange=s.getRange(2,22,MB.BOOKIE_HEALTH_ROWS,1), bookmakerValues=bookmakerRange.getDisplayValues();
  bookmakerRange.setValues(bookmakerValues.map(r=>[baseBookmakerName_(r[0])]));
  s.getRange('P1:T1').setBackground('#183153').setFontColor('white').setFontWeight('bold');
  s.getRange('V1').setBackground('#183153').setFontColor('white').setFontWeight('bold');
}

function setupEntrySheet_(s, headers) {
  mbEnsureSize_(s, MB.ROWS, headers.length);
  s.getRange(1,1,1,headers.length).setValues([headers]);
  s.setFrozenRows(1);
  const filter = s.getFilter();
  if (!filter) s.getRange(1,1,s.getMaxRows(),headers.length).createFilter();
}
