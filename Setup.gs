function buildWorkbook() {
  MB.SHEETS.forEach(n => mbSheet_(n));
  setupSettings_();
  setupPromos_();
  setupEntrySheet_(mbSheet_('BTO'), MB.BTO_HEADERS);
  setupEntrySheet_(mbSheet_('NonPromos'), MB.NON_HEADERS);
  seedPromoExamples_();
  refreshFormulas();
  refreshDropdowns();
  refreshDashboard();
  formatWorkbook_();
  runAudit();
  SpreadsheetApp.getActive().toast('Matched-betting workbook is ready.');
}

function seedPromoExamples_() {
  const s=mbSheet_('Promos');
  if (s.getRange('A6:A').getValues().some(r=>r[0]!=='')) return;
  const today=new Date();
  s.getRange('A6:G8').setValues([
    [today,'','SGM3 Leg','Starter example: edit or delete',100,2.10,'W'],
    [today,'','Free Hit Single','Starter example: edit or delete',50,4.00,'B'],
    [today,'','SRM','Starter example: edit or delete',25,3.20,'L']]);
  s.getRange('I6:J8').setValues([[30,0],[-50,50],[-25,0]]);
}

/** Migrates the original 16-column Promo table without changing its entries. */
function setupPromos_() {
  const s=mbSheet_('Promos'), oldHeaders=s.getRange(1,1,1,Math.min(16,s.getMaxColumns())).getValues()[0];
  const isOld=oldHeaders.indexOf('Entry ID')>=0 || oldHeaders.indexOf('Estimated EV')>=0;
  let entries=[];
  if (isOld) {
    const index={}; oldHeaders.forEach((h,i)=>index[h]=i);
    const values=s.getRange(2,1,Math.max(1,s.getLastRow()-1),oldHeaders.length).getValues();
    entries=values.filter(r=>r[index['Date']]!=='').map(r=>[
      r[index['Date']],r[index['Account']],r[index['Promo Type']],r[index['Notes']],r[index['Turnover']],
      r[index['Odds']],r[index['Result']],'',r[index['Cash Change']],r[index['Bonus Change']]
    ]);
    if (s.getFilter()) s.getFilter().remove();
    s.clear();
  }
  mbEnsureSize_(s,MB.ROWS,MB.PROMO_HEADERS.length);
  if (s.getMaxColumns()>MB.PROMO_HEADERS.length) s.deleteColumns(MB.PROMO_HEADERS.length+1,s.getMaxColumns()-MB.PROMO_HEADERS.length);
  s.getRange(MB.PROMO_HEADER_ROW,1,1,MB.PROMO_HEADERS.length).setValues([MB.PROMO_HEADERS]);
  s.getRange('A1:C1').setValues([['Total Turnover','Total EV Taken','Total EV %']]);
  if (entries.length) s.getRange(MB.PROMO_DATA_ROW,1,entries.length,entries[0].length).setValues(entries);
  s.setFrozenRows(MB.PROMO_HEADER_ROW);
  if (!s.getFilter()) s.getRange(MB.PROMO_HEADER_ROW,1,s.getMaxRows()-MB.PROMO_HEADER_ROW+1,MB.PROMO_HEADERS.length).createFilter();
}

function setupSettings_() {
  const s = mbSheet_('Settings');
  mbEnsureSize_(s, 100, 14);
  const headings = [['Promo Type','EV %','','EV Tier','Minimum %','Maximum %','','Bookmaker Accounts','BTO Methods','Non-Promo Types','Results','Account Statuses','Transaction Types','Banks']];
  s.getRange(1,1,1,14).setValues(headings);
  if (!s.getRange('A2').getValue()) s.getRange(2,1,MB.PROMOS.length,2).setValues(MB.PROMOS);
  if (!s.getRange('D2').getValue()) s.getRange('D2:F7').setValues([['A+',.31,1],['A',.20,.3099],['B',.15,.1999],['C',.10,.1499],['D',.05,.0999],['E',0,.0499]]);
  if (!s.getRange('I2').getValue()) s.getRange(2,9,MB.BTO_METHODS.length,1).setValues(MB.BTO_METHODS.map(x=>[x]));
  if (!s.getRange('J2').getValue()) s.getRange(2,10,MB.NON_TYPES.length,1).setValues(MB.NON_TYPES.map(x=>[x]));
  if (!s.getRange('K2').getValue()) s.getRange(2,11,MB.RESULTS.length,1).setValues(MB.RESULTS.map(x=>[x]));
  if (!s.getRange('L2').getValue()) s.getRange('L2:L4').setValues([['Open'],['Restricted'],['Closed']]);
  if (!s.getRange('M2').getValue()) s.getRange('M2:M4').setValues([['Deposit'],['Withdrawal'],['Adjustment']]);
  s.getRange('D10').setValue('Default BTO rate when BSP is missing');
  if (s.getRange('E10').isBlank()) s.getRange('E10').setValue(.75);
  s.getRange('A:B').setNumberFormat('0.00%'); s.getRange('A:A').setNumberFormat('@');
  s.getRange('B2:B100').setNumberFormat('0.00%'); s.getRange('E2:F10').setNumberFormat('0.00%');
}

function setupEntrySheet_(s, headers) {
  mbEnsureSize_(s, MB.ROWS, headers.length);
  s.getRange(1,1,1,headers.length).setValues([headers]);
  s.setFrozenRows(1);
  const filter = s.getFilter();
  if (!filter) s.getRange(1,1,s.getMaxRows(),headers.length).createFilter();
}
