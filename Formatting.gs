function formatWorkbook_() {
  const navy='#1f2933',green='#17652a',gold='#9a7600',red='#bf1e24',money='[$$-en-AU]#,##0.00;[Red]-[$$-en-AU]#,##0.00',pct='0.00%;[Red]-0.00%',decimal='0.00';
  MB.SHEETS.forEach(n=>mbSheet_(n).getDataRange().setFontFamily('Arial').setVerticalAlignment('middle'));
  const specs={Promos:{money:['Turnover','Estimated EV','Cash Change','Bonus Change'],decimal:['Odds'],pct:[]},BTO:{money:['Turnover','EV Taken','Bookie Change','Betfair Lay','Betfair Change'],decimal:['Odds','BSP','Lay Odds'],pct:['Commission']},NonPromos:{money:['Turnover','Expected EV','Bookie Change','Betfair Lay','Betfair Change'],decimal:['Odds','BSP','Lay Odds'],pct:['EV %','Commission']}};
  Object.keys(specs).forEach(name=>{const s=mbSheet_(name);s.getRange(1,1,1,s.getLastColumn()).setBackground(navy).setFontColor('white').setFontWeight('bold').setWrap(true);s.getRange(2,mbColumn_(s,'Date'),MB.ROWS-1,1).setNumberFormat('dd/mm/yyyy');specs[name].money.forEach(h=>s.getRange(2,mbColumn_(s,h),MB.ROWS-1,1).setNumberFormat(money));specs[name].decimal.forEach(h=>s.getRange(2,mbColumn_(s,h),MB.ROWS-1,1).setNumberFormat(decimal));specs[name].pct.forEach(h=>s.getRange(2,mbColumn_(s,h),MB.ROWS-1,1).setNumberFormat(pct));s.setColumnWidth(mbColumn_(s,'Notes'),260);});
  protectFormulaColumns_(mbSheet_('Promos'),['Estimated EV','Cash Change','Bonus Change']);protectFormulaColumns_(mbSheet_('BTO'),['EV Taken','Bookie Change','Betfair Change']);protectFormulaColumns_(mbSheet_('NonPromos'),['Expected EV','EV %','Bookie Change','Betfair Change']);
  const mm=mbSheet_('Monthly Metrics');mm.getRange('A2:A500').setNumberFormat('MMM yyyy');[2,3,5,7,8,10,12,13,15,17,18,19,20].forEach(c=>mm.getRange(2,c,499,1).setNumberFormat(money));[4,6,9,11,14,16].forEach(c=>mm.getRange(2,c,499,1).setNumberFormat(pct));
  const ac=mbSheet_('Accounts');ac.getRange(1,1,1,12).setBackground('#4c1d95').setFontColor('white').setFontWeight('bold');ac.getRange(4,1,1,12).setBackground('#5b21b6').setFontColor('white').setFontWeight('bold');ac.getRange(5,3,ac.getMaxRows()-4,9).setNumberFormat(money);
  [['EV Breakdown',green,gold,red],['Master PnL',green,gold,red]].forEach(x=>mbSheet_(x[0]).getDataRange().setFontFamily('Arial'));
  ['Promos','BTO','NonPromos','Monthly Metrics','EV Breakdown','Master PnL','Accounts'].forEach(n=>{const s=mbSheet_(n),range=s.getDataRange(),rules=s.getConditionalFormatRules().filter(r=>!r.getBooleanCondition());rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setFontColor('#c00000').setRanges([range]).build());s.setConditionalFormatRules(rules);});
}

function protectFormulaColumns_(sheet, headers) {
  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).filter(p=>p.getDescription()==='Managed formula column').forEach(p=>p.remove());
  headers.forEach(h=>sheet.getRange(2,mbColumn_(sheet,h),MB.ROWS-1,1).protect().setDescription('Managed formula column').setWarningOnly(true));
}
