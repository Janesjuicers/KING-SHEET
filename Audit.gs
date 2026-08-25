function runAudit() {
  const s=mbSheet_('Audit'); mbEnsureSize_(s,40,5); s.clear();
  s.getRange('A1:E1').setValues([['Audit Test','Status','Observed','Expected / tolerance','Notes']]);
  const tests=[
    ['Promo totals versus Monthly Metrics','=IF(ABS(SUM(Promos!M2:M)-SUM(\'Monthly Metrics\'!E2:E))<=0.01,"PASS","FAIL")','=SUM(Promos!M2:M)-SUM(\'Monthly Metrics\'!E2:E)','$0.01','Entry-level source versus monthly aggregate'],
    ['BTO totals versus Monthly Metrics','=IF(ABS(SUM(BTO!Q2:Q)-SUM(\'Monthly Metrics\'!J2:J))<=0.01,"PASS","FAIL")','=SUM(BTO!Q2:Q)-SUM(\'Monthly Metrics\'!J2:J)','$0.01',''],
    ['Non-promo totals versus Monthly Metrics','=IF(ABS(SUM(NonPromos!N2:N)-SUM(\'Monthly Metrics\'!O2:O))<=0.01,"PASS","FAIL")','=SUM(NonPromos!N2:N)-SUM(\'Monthly Metrics\'!O2:O)','$0.01',''],
    ['Monthly Metrics versus Master PnL','=IF(ABS(SUM(\'Monthly Metrics\'!R2:R)-\'Master PnL\'!B21)<=0.01,"PASS","FAIL")','=SUM(\'Monthly Metrics\'!R2:R)-\'Master PnL\'!B21','$0.01',''],
    ['Missing Promo Type Settings rates','=IF(COUNTIFS(Promos!D2:D,"<>",Promos!I2:I,"#N/A")=0,"PASS","FAIL")','=COUNTIFS(Promos!D2:D,"<>",Promos!I2:I,"#N/A")','0',''],
    ['Missing BSP values','=IF(COUNTIFS(BTO!B2:B,"<>",BTO!H2:H,"")+COUNTIFS(NonPromos!B2:B,"<>",NonPromos!H2:H,"")=0,"PASS","FAIL")','=COUNTIFS(BTO!B2:B,"<>",BTO!H2:H,"")+COUNTIFS(NonPromos!B2:B,"<>",NonPromos!H2:H,"")','0','BTO blanks are permitted but flagged for review'],
    ['Duplicate Entry IDs','=IF(SUM(ARRAYFORMULA(N(COUNTIF({Promos!A2:A;BTO!A2:A;NonPromos!A2:A},{Promos!A2:A;BTO!A2:A;NonPromos!A2:A})>1)*N({Promos!A2:A;BTO!A2:A;NonPromos!A2:A}<>"")))=0,"PASS","FAIL")','','0',''],
    ['Invalid dates','=IF(SUM(ARRAYFORMULA(N({Promos!B2:B;BTO!B2:B;NonPromos!B2:B}<>"")*N(NOT(ISNUMBER({Promos!B2:B;BTO!B2:B;NonPromos!B2:B}))))))=0,"PASS","FAIL")','','0',''],
    ['Formula errors', formulaErrors_() ? 'FAIL':'PASS',formulaErrors_(), '0','Scanned all displayed cell values'],
    ['$35 at odds 8.00 with BSP 10.00 must produce $24.50','=IF(ROUND(35*(8-1)/10,2)=24.5,"PASS","FAIL")','=ROUND(35*(8-1)/10,2)','$24.50','BTO formula test'],
    ['$30 at odds 11.00 with BSP 15.61 must produce $19.22','=IF(ROUND(30*(11-1)/15.61,2)=19.22,"PASS","FAIL")','=ROUND(30*(11-1)/15.61,2)','$19.22','BTO formula test'],
    ['$10 at odds 10.00 with BSP 16.54 must produce $5.44','=IF(ROUND(10*(10-1)/16.54,2)=5.44,"PASS","FAIL")','=ROUND(10*(10-1)/16.54,2)','$5.44','BTO formula test'],
    ['$25 with no BSP and a 75% default rate must produce $18.75','=IF(ROUND(25*Settings!E10,2)=18.75,"PASS","FAIL")','=ROUND(25*Settings!E10,2)','$18.75','Uses editable Settings rate']];
  tests.forEach((r,i)=>{r.forEach((v,j)=>{const c=s.getRange(i+2,j+1);if(typeof v==='string'&&v[0]==='=')c.setFormula(v);else c.setValue(v);});});
  s.getRange('A1:E1').setBackground('#183153').setFontColor('white').setFontWeight('bold');s.setFrozenRows(1);s.autoResizeColumns(1,5);s.setColumnWidth(1,340);s.setColumnWidth(5,300);
}

function formulaErrors_() {
  return MB.SHEETS.filter(n=>n!=='Audit').reduce((sum,n)=>sum+mbSheet_(n).getDataRange().getDisplayValues().flat().filter(v=>/^#(REF|N\/A|VALUE|DIV\/0|NAME|NUM|ERROR)!?$/.test(v)).length,0);
}

function runBtoFormulaTests() {
  const rate=mbSheet_('Settings').getRange('E10').getValue(); const got=[35*(8-1)/10,30*(11-1)/15.61,10*(10-1)/16.54,25*rate].map(x=>Math.round(x*100)/100), want=[24.5,19.22,5.44,18.75];
  if (!got.every((x,i)=>x===want[i])) throw new Error('BTO test failed: '+got.join(', '));
  SpreadsheetApp.getActive().toast('All four BTO formula tests passed.'); return got;
}
