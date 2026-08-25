function formatWorkbook_() {
  const navy='#183153', blue='#d9eaf7', green='#d9ead3', red='#f4cccc';
  MB.SHEETS.forEach(n=>{const s=mbSheet_(n);s.setFrozenRows(n==='EV Breakdown'?7:1);s.getDataRange().setFontFamily('Arial').setVerticalAlignment('middle');});
  ['Promos','BTO','NonPromos','Monthly Metrics'].forEach(n=>{const s=mbSheet_(n);const cols=s.getLastColumn();s.getRange(1,1,1,cols).setBackground(navy).setFontColor('white').setFontWeight('bold').setWrap(true);s.setRowHeight(1,42);s.getRange(2,1,MB.ROWS-1,1).setNumberFormat('dd/mm/yyyy');if(n==='NonPromos')s.getRange(2,15,MB.ROWS-1,1).setNumberFormat('MMM yyyy');});
  mbSheet_('Settings').getRange('A1:N1').setBackground(navy).setFontColor('white').setFontWeight('bold');
  const money='$#,##0.00;[Red]-$#,##0.00', pct='0.00%;[Red]-0.00%';
  const formats={Promos:{money:[5,8,9,10],pct:[]},BTO:{money:[5,9,11,12,15,16],pct:[10,14,17]},NonPromos:{money:[5,9,11,12,13],pct:[10,14]}};
  Object.keys(formats).forEach(n=>{let s=mbSheet_(n);formats[n].money.forEach(c=>s.getRange(2,c,MB.ROWS-1,1).setNumberFormat(money));formats[n].pct.forEach(c=>s.getRange(2,c,MB.ROWS-1,1).setNumberFormat(pct));s.autoResizeColumns(1,s.getLastColumn());s.setColumnWidth(5,220);});
  const mm=mbSheet_('Monthly Metrics');mm.getRange('A2:A500').setNumberFormat('MMM yyyy');[2,3,5,7,8,10,12,13,15,17,18,19,20].forEach(c=>mm.getRange(2,c,499,1).setNumberFormat(money));[4,6,9,11,14,16].forEach(c=>mm.getRange(2,c,499,1).setNumberFormat(pct));
  [['Promos',[8,9,10]],['BTO',[9,10,16,17]],['NonPromos',[9,10,13,14,15]]].forEach(x=>protectFormulaColumns_(mbSheet_(x[0]),x[1]));
  ['Promos','BTO','NonPromos','Monthly Metrics','EV Breakdown','Master PnL'].forEach(n=>{let s=mbSheet_(n);let range=s.getDataRange();let rules=s.getConditionalFormatRules().filter(r=>!r.getBooleanCondition());rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setFontColor('#b91c1c').setBackground(red).setRanges([range]).build());s.setConditionalFormatRules(rules);});
}

function protectFormulaColumns_(sheet, cols) {
  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).filter(p=>p.getDescription()==='Managed formula column').forEach(p=>p.remove());
  cols.forEach(c=>sheet.getRange(2,c,MB.ROWS-1,1).protect().setDescription('Managed formula column').setWarningOnly(true));
}
