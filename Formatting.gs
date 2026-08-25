function formatWorkbook_() {
  const navy='#183153', blue='#d9eaf7', green='#d9ead3', red='#f4cccc';
  MB.SHEETS.forEach(n=>{const s=mbSheet_(n);s.setFrozenRows(n==='EV Breakdown'?7:n==='Promos'?MB.PROMO_HEADER_ROW:1);s.getDataRange().setFontFamily('Arial').setVerticalAlignment('middle');});
  ['BTO','NonPromos','Monthly Metrics'].forEach(n=>{const s=mbSheet_(n);const cols=s.getLastColumn();s.getRange(1,1,1,cols).setBackground(navy).setFontColor('white').setFontWeight('bold').setWrap(true);s.setRowHeight(1,42);s.getRange(2,2,MB.ROWS-1,1).setNumberFormat('dd/mm/yyyy');const month=n==='BTO'?19:16;s.getRange(2,month,MB.ROWS-1,1).setNumberFormat('MMM yyyy');});
  mbSheet_('Settings').getRange('A1:N1').setBackground(navy).setFontColor('white').setFontWeight('bold');
  const money='$#,##0.00;[Red]-$#,##0.00', pct='0.00%;[Red]-0.00%';
  const promos=mbSheet_('Promos');
  promos.getRange('A1:C1').setBackground('#d9ead3').setFontWeight('bold');promos.getRange('A2:B2').setNumberFormat(money);promos.getRange('C2').setNumberFormat(pct);
  promos.getRange(5,1,1,10).setBackground('#6aa84f').setFontColor('white').setFontWeight('bold').setWrap(true);promos.setRowHeight(5,42);
  promos.getRange(6,1,MB.ROWS-5,1).setNumberFormat('dd/mm/yyyy');[5,6,8,9,10].forEach(c=>promos.getRange(6,c,MB.ROWS-5,1).setNumberFormat(money));
  promos.getRange(6,8,MB.ROWS-5,1).setBackground('#d9d9d9');promos.autoResizeColumns(1,10);promos.setColumnWidth(4,220);
  const formats={BTO:{money:[6,10,12,13,16,17],pct:[11,15,18]},NonPromos:{money:[6,10,12,13,14],pct:[11,15]}};
  Object.keys(formats).forEach(n=>{let s=mbSheet_(n);formats[n].money.forEach(c=>s.getRange(2,c,MB.ROWS-1,1).setNumberFormat(money));formats[n].pct.forEach(c=>s.getRange(2,c,MB.ROWS-1,1).setNumberFormat(pct));s.autoResizeColumns(1,s.getLastColumn());s.setColumnWidth(5,220);});
  const mm=mbSheet_('Monthly Metrics');mm.getRange('A2:A500').setNumberFormat('MMM yyyy');[2,3,5,7,8,10,12,13,15,17,18,19,20].forEach(c=>mm.getRange(2,c,499,1).setNumberFormat(money));[4,6,9,11,14,16].forEach(c=>mm.getRange(2,c,499,1).setNumberFormat(pct));
  protectFormulaColumns_(promos,[8]);
  [['BTO',[1,10,11,17,18,19]],['NonPromos',[1,10,11,14,15,16]]].forEach(x=>protectFormulaColumns_(mbSheet_(x[0]),x[1]));
  ['Promos','BTO','NonPromos','Monthly Metrics','EV Breakdown','Master PnL'].forEach(n=>{let s=mbSheet_(n);let range=s.getDataRange();let rules=s.getConditionalFormatRules().filter(r=>!r.getBooleanCondition());rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setFontColor('#b91c1c').setBackground(red).setRanges([range]).build());s.setConditionalFormatRules(rules);});
}

function protectFormulaColumns_(sheet, cols) {
  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).filter(p=>p.getDescription()==='Managed formula column').forEach(p=>p.remove());
  cols.forEach(c=>sheet.getRange(2,c,MB.ROWS-1,1).protect().setDescription('Managed formula column').setWarningOnly(true));
}
