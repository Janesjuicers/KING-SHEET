function refreshDashboard() {
  buildBreakdown_(); buildMaster_(); buildCharts_();
}

function periodBounds_() {
  return {start:'IF($B$2="Custom",$B$3,IF($B$2="All Time",DATE(1900,1,1),IF($B$2="Last 30 Days",TODAY()-29,EDATE(TODAY(),-VALUE(REGEXEXTRACT($B$2,"\\d+"))+1))))', end:'IF($B$2="Custom",$B$4,TODAY())'};
}

function buildBreakdown_() {
  const s=mbSheet_('EV Breakdown'); mbEnsureSize_(s,100,20);
  s.getRange('A1:B4').setValues([['Reporting Period','All Time'],['Period','All Time'],['From Date',''],['To Date','']]);
  if (MB.PERIODS.indexOf(s.getRange('B2').getValue())<0) s.getRange('B2').setValue('All Time');
  const b=periodBounds_();
  breakdownTable_(s,6,1,'Promo breakdown',['Promo Type','Turnover','Estimated EV','Actual PnL','Average EV %','POT %','% Total Promo Turnover'],MB.PROMOS.map(x=>x[0]),'Promos','D','F','J','M',b,true);
  breakdownTable_(s,6,9,'BTO breakdown',['BTO Method','Turnover','EV Taken','EV %','Actual Return','Actual BTO %'],MB.BTO_METHODS,'BTO','C','E','I','P',b,false,'A');
  breakdownTable_(s,6,16,'Non-promo breakdown',['Non-Promo Type','Turnover','Expected EV','EV %','Actual PnL','POT %','% Total Non-Promo Turnover'],MB.NON_TYPES,'NonPromos','D','F','J','N',b,true);
}

function breakdownTable_(s,row,col,title,headers,items,source,typeCol,toCol,evCol,pnlCol,b,share,dateCol) {
  s.getRange(row,col).setValue(title); s.getRange(row+1,col,1,headers.length).setValues([headers]);
  s.getRange(row+2,col,items.length,1).setValues(items.map(x=>[x]));
  items.forEach((_,i)=>{
    const r=row+2+i, key=s.getRange(r,col).getA1Notation(), q=n=>"'"+source+"'!"+n+":"+n;
    const dates=dateCol||'B', sum=(field)=>`=SUMIFS(${q(field)},${q(typeCol)},${key},${q(dates)},">="&${b.start},${q(dates)},"<="&${b.end})`;
    s.getRange(r,col+1).setFormula(sum(toCol)); s.getRange(r,col+2).setFormula(sum(evCol));
    if (source==='Promos') { s.getRange(r,col+3).setFormula(sum(pnlCol)); s.getRange(r,col+4).setFormula(`=IF(${s.getRange(r,col+1).getA1Notation()}=0,"",${s.getRange(r,col+2).getA1Notation()}/${s.getRange(r,col+1).getA1Notation()})`); s.getRange(r,col+5).setFormula(`=IF(${s.getRange(r,col+1).getA1Notation()}=0,"",${s.getRange(r,col+3).getA1Notation()}/${s.getRange(r,col+1).getA1Notation()})`); }
    else { s.getRange(r,col+3).setFormula(`=IF(${s.getRange(r,col+1).getA1Notation()}=0,"",${s.getRange(r,col+2).getA1Notation()}/${s.getRange(r,col+1).getA1Notation()})`); s.getRange(r,col+4).setFormula(sum(pnlCol)); s.getRange(r,col+5).setFormula(`=IF(${s.getRange(r,col+1).getA1Notation()}=0,"",${s.getRange(r,col+4).getA1Notation()}/${s.getRange(r,col+1).getA1Notation()})`); }
    if (share) s.getRange(r,col+6).setFormula(`=IF(${s.getRange(r,col+1).getA1Notation()}=0,"",${s.getRange(r,col+1).getA1Notation()}/SUM(${s.getRange(row+2,col+1,items.length,1).getA1Notation()}))`);
  });
}

function buildMaster_() {
  const s=mbSheet_('Master PnL'); mbEnsureSize_(s,80,20); s.getRange('A1').setValue('MASTER PnL DASHBOARD');
  const rows=[['PROMOS',''],['Total Promo Turnover','=SUM(Promos!F2:F)'],['Total Promo Estimated EV','=SUM(Promos!J2:J)'],['Promo EV %','=IF(B3=0,"",B4/B3)'],['Total Promo Actual PnL','=SUM(Promos!M2:M)'],['Promo POT %','=IF(B3=0,"",B6/B3)'],['BTO',''],['Total BTO Turnover','=SUM(BTO!E2:E)'],['Total EV Taken','=SUM(BTO!I2:I)'],['Expected BTO %','=IF(B9=0,"",B10/B9)'],['Total Actual Return','=SUM(BTO!P2:P)'],['Actual BTO %','=IF(B9=0,"",B12/B9)'],['NON-PROMOS',''],['Total Non-Promo Turnover','=SUM(NonPromos!F2:F)'],['Total Non-Promo EV','=SUM(NonPromos!J2:J)'],['Non-Promo EV %','=IF(B15=0,"",B16/B15)'],['Total Non-Promo Actual PnL','=SUM(NonPromos!N2:N)'],['Non-Promo POT %','=IF(B15=0,"",B18/B15)'],['OVERALL',''],['Overall EV','=B4+B10+B16'],['Overall Actual PnL','=B6+B12+B18'],['Actual less EV','=B21-B20'],['Cumulative EV','=B20'],['Cumulative actual PnL','=B21']];
  s.getRange(2,1,rows.length,1).setValues(rows.map(x=>[x[0]])); rows.forEach((x,i)=>{if(x[1])s.getRange(2+i,2).setFormula(x[1]);});
  s.getRange('D2:H2').setValues([['Results by Promo Type','Turnover','Estimated EV','Actual PnL','POT %']]);
  s.getRange(3,4,MB.PROMOS.length,1).setValues(MB.PROMOS.map(x=>[x[0]]));
  MB.PROMOS.forEach((_,i)=>{let r=3+i;s.getRange(r,5).setFormula(`=SUMIF(Promos!D:D,D${r},Promos!F:F)`);s.getRange(r,6).setFormula(`=SUMIF(Promos!D:D,D${r},Promos!J:J)`);s.getRange(r,7).setFormula(`=SUMIF(Promos!D:D,D${r},Promos!M:M)`);s.getRange(r,8).setFormula(`=IF(E${r}=0,"",G${r}/E${r})`);});
  s.getRange('J2:M2').setValues([['Results by EV Tier','Turnover','Estimated EV','Actual PnL']]); const tiers=['A+','A','B','C','D','E']; s.getRange(3,10,6,1).setValues(tiers.map(x=>[x])); tiers.forEach((_,i)=>{let r=3+i;s.getRange(r,11).setFormula(`=SUMIF(Promos!O:O,J${r},Promos!F:F)`);s.getRange(r,12).setFormula(`=SUMIF(Promos!O:O,J${r},Promos!J:J)`);s.getRange(r,13).setFormula(`=SUMIF(Promos!O:O,J${r},Promos!M:M)`);});
}
