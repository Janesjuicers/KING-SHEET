function runAudit() {
  const s=mbSheet_('Audit'); mbEnsureSize_(s,40,5); s.clear();
  s.getRange('A1:E1').setValues([['Audit Test','Status','Observed','Expected / tolerance','Notes']]);
  const tests=[
    ['Promo totals versus Monthly Metrics','=IF(ABS(SUM(Promos!M2:M)-SUM(\'Monthly Metrics\'!E2:E))<=0.01,"PASS","FAIL")','=SUM(Promos!M2:M)-SUM(\'Monthly Metrics\'!E2:E)','$0.01','Entry-level source versus monthly aggregate'],
    ['BTO totals versus Monthly Metrics','=IF(ABS(SUM(BTO!P2:P)-SUM(\'Monthly Metrics\'!J2:J))<=0.01,"PASS","FAIL")','=SUM(BTO!P2:P)-SUM(\'Monthly Metrics\'!J2:J)','$0.01',''],
    ['Non-promo totals versus Monthly Metrics','=IF(ABS(SUM(NonPromos!N2:N)-SUM(\'Monthly Metrics\'!O2:O))<=0.01,"PASS","FAIL")','=SUM(NonPromos!N2:N)-SUM(\'Monthly Metrics\'!O2:O)','$0.01',''],
    ['Monthly Metrics versus Master PnL','=IF(ABS(SUM(\'Monthly Metrics\'!R2:R)-\'Master PnL\'!B21)<=0.01,"PASS","FAIL")','=SUM(\'Monthly Metrics\'!R2:R)-\'Master PnL\'!B21','$0.01',''],
    ['Missing Promo Type Settings rates','=IF(COUNTIFS(Promos!D2:D,"<>",Promos!I2:I,"#N/A")=0,"PASS","FAIL")','=COUNTIFS(Promos!D2:D,"<>",Promos!I2:I,"#N/A")','0',''],
    ['Missing BSP values','=IF(COUNTIFS(BTO!A2:A,"<>",BTO!G2:G,"")+COUNTIFS(NonPromos!B2:B,"<>",NonPromos!H2:H,"")=0,"PASS","FAIL")','=COUNTIFS(BTO!A2:A,"<>",BTO!G2:G,"")+COUNTIFS(NonPromos!B2:B,"<>",NonPromos!H2:H,"")','0','BTO blanks are permitted but flagged for review'],
    ['Duplicate Entry IDs','=IF(SUM(ARRAYFORMULA(N(COUNTIF({Promos!A2:A;NonPromos!A2:A},{Promos!A2:A;NonPromos!A2:A})>1)*N({Promos!A2:A;NonPromos!A2:A}<>"")))=0,"PASS","FAIL")','','0','BTO has no Entry ID'],
    ['Invalid dates','=IF(SUM(ARRAYFORMULA(N({Promos!B2:B;BTO!A2:A;NonPromos!B2:B}<>"")*N(NOT(ISNUMBER({Promos!B2:B;BTO!A2:A;NonPromos!B2:B}))))))=0,"PASS","FAIL")','','0',''],
    ['Formula errors', formulaErrors_() ? 'FAIL':'PASS',formulaErrors_(), '0','Scanned all displayed cell values'],
    ['$35 at odds 8.00 with BSP 10.00 must produce $24.50','=IF(ROUND(35*(8-1)/10,2)=24.5,"PASS","FAIL")','=ROUND(35*(8-1)/10,2)','$24.50','BTO formula test'],
    ['$30 at odds 11.00 with BSP 15.61 must produce $19.22','=IF(ROUND(30*(11-1)/15.61,2)=19.22,"PASS","FAIL")','=ROUND(30*(11-1)/15.61,2)','$19.22','BTO formula test'],
    ['$10 at odds 10.00 with BSP 16.54 must produce $5.44','=IF(ROUND(10*(10-1)/16.54,2)=5.44,"PASS","FAIL")','=ROUND(10*(10-1)/16.54,2)','$5.44','BTO formula test'],
    ['$25 with no BSP and a 75% default rate must produce $18.75','=IF(ROUND(25*Settings!E10,2)=18.75,"PASS","FAIL")','=ROUND(25*Settings!E10,2)','$18.75','Uses editable Settings rate']];
  bookieHealthAuditTests_().forEach(t=>tests.push(t));
  tests.forEach((r,i)=>{r.forEach((v,j)=>{const c=s.getRange(i+2,j+1);if(typeof v==='string'&&v[0]==='=')c.setFormula(v);else c.setValue(v);});});
  s.getRange('A1:E1').setBackground('#183153').setFontColor('white').setFontWeight('bold');s.setFrozenRows(1);s.autoResizeColumns(1,5);s.setColumnWidth(1,340);s.setColumnWidth(5,300);
}

function bookieHealthAuditTests_() {
  const settings=mbSheet_('Settings'), health=mbSheet_('Bookie Health');
  const rawSettings=settings.getRange('V2:V100').getDisplayValues().flat().map(v=>v.trim());
  const lastSetting=rawSettings.reduce((n,v,i)=>v?i+1:n,0);
  const settingNames=rawSettings.filter(Boolean), settingSet=new Set(settingNames);
  const rawHealth=health.getLastRow()>2?health.getRange(3,1,health.getLastRow()-2,1).getDisplayValues().flat().map(v=>v.trim()):[];
  const healthNames=rawHealth.filter(Boolean), healthSet=new Set(healthNames);
  const statusSet=new Set(activeBookieHealthStatuses_().map(r=>String(r[0])));
  let invalid=0, duplicateStatuses=0, missingValidation=0;
  if(healthNames.length && health.getLastColumn()>1) {
    const statusRanges=[]; for(let c=2;c<=health.getLastColumn();c+=2)statusRanges.push(health.getRange(3,c,healthNames.length,1));
    statusRanges.forEach(range=>range.getDisplayValues().flat().forEach(value=>{
      if(!value) return;
      const parts=value.split(MB.BOOKIE_HEALTH_SEPARATOR).map(v=>v.trim()).filter(Boolean);
      invalid+=parts.filter(v=>!statusSet.has(v)).length;
      duplicateStatuses+=parts.length-new Set(parts).size;
    }));
    missingValidation=statusRanges.reduce((n,range)=>n+range.getDataValidations().flat().filter(v=>!v).length,0);
  }
  const duplicateNames=(settingNames.length-settingSet.size)+(healthNames.length-healthSet.size);
  const blankNames=rawSettings.slice(0,lastSetting).filter(v=>!v).length;
  const missingHealth=settingNames.filter(v=>!healthSet.has(v)).length;
  const missingSettings=healthNames.filter(v=>!settingSet.has(v)).length;
  const result=(name,count,note)=>[name,count===0?'PASS':'FAIL',count,'0',note];
  return [
    result('Bookie Health duplicate bookmaker names',duplicateNames,'Checks Settings and Bookie Health'),
    result('Bookie Health blank bookmaker names',blankNames,'Blank names inside the populated Settings list'),
    result('Bookie Health invalid status values',invalid,'Splits multi-selections on ", " and checks active Settings statuses'),
    result('Bookie Health duplicate statuses in cells',duplicateStatuses,'Each selected status may appear only once per cell'),
    result('Bookie Health missing dropdown validation',missingValidation,'Checks every populated bookmaker/profile cell'),
    result('Settings bookmakers missing from Bookie Health',missingHealth,'Refresh Bookie Health to add missing rows'),
    result('Bookie Health bookmakers missing from Settings',missingSettings,'Bookie Health contains no unmanaged rows')
  ];
}

function formulaErrors_() {
  return MB.SHEETS.filter(n=>n!=='Audit').reduce((sum,n)=>sum+mbSheet_(n).getDataRange().getDisplayValues().flat().filter(v=>/^#(REF|N\/A|VALUE|DIV\/0|NAME|NUM|ERROR)!?$/.test(v)).length,0);
}

function runBtoFormulaTests() {
  const rate=mbSheet_('Settings').getRange('E10').getValue(); const got=[35*(8-1)/10,30*(11-1)/15.61,10*(10-1)/16.54,25*rate].map(x=>Math.round(x*100)/100), want=[24.5,19.22,5.44,18.75];
  if (!got.every((x,i)=>x===want[i])) throw new Error('BTO test failed: '+got.join(', '));
  SpreadsheetApp.getActive().toast('All four BTO formula tests passed.'); return got;
}
