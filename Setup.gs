function buildWorkbook() {
  MB.SHEETS.forEach(n => mbSheet_(n));
  setupSettings_();
  refreshBookieHealth();
  setupEntrySheet_(mbSheet_('Promos'), MB.PROMO_HEADERS);
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
  if (s.getRange('B2:B').getValues().some(r=>r[0]!=='')) return;
  const today=new Date();
  s.getRange('B2:H4').setValues([
    [today,'','SGM3 Leg','Starter example: edit or delete',100,2.10,'W'],
    [today,'','Free Hit Single','Starter example: edit or delete',50,4.00,'B'],
    [today,'','SRM','Starter example: edit or delete',25,3.20,'L']]);
  s.getRange('K2:L4').setValues([[30,0],[-50,50],[-25,0]]);
}

function setupSettings_() {
  const s = mbSheet_('Settings');
  mbEnsureSize_(s, 100, 22);
  const headings = [['Promo Type','EV %','','EV Tier','Minimum %','Maximum %','','Bookmaker Accounts','BTO Methods','Non-Promo Types','Results','Account Statuses','Transaction Types','Banks']];
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
    ['Promos','BTO','NonPromos'].forEach(n=>mbSheet_(n).getRange(2,3,MB.ROWS-1,1).getDisplayValues().forEach(r=>candidates.push(r[0])));
    const names=[...new Set(candidates.map(v=>String(v).trim()).filter(Boolean))];
    if (names.length) s.getRange(2,22,names.length,1).setValues(names.map(v=>[v]));
  }
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
