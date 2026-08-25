function buildCharts_() {
  const mm=mbSheet_('Monthly Metrics'), master=mbSheet_('Master PnL');
  [mm,master].forEach(s=>s.getCharts().forEach(c=>s.removeChart(c)));
  mm.insertChart(mm.newChart().asLineChart().addRange(mm.getRange('A1:A200')).addRange(mm.getRange('S1:T200')).setPosition(24,2,0,0).setOption('title','Cumulative EV vs Cumulative actual PnL').build());
  const specs=[
    [mm.getRange('A1:A200'),mm.getRange('Q1:R200'),'Monthly EV versus actual results',2,15,'ColumnChart'],
    [mm.getRange('A1:A200'),mm.getRange('S1:T200'),'Cumulative EV versus cumulative PnL',20,15,'LineChart'],
    [master.getRange('J2:K8'),null,'Promo turnover by EV tier',38,15,'PieChart'],
    [master.getRange('D2:D22'),master.getRange('F2:G22'),'Promo estimated EV versus actual PnL by Promo Type',2,23,'ColumnChart'],
    [mm.getRange('A1:A200'),mm.getRange('I1:I200'),'Monthly BTO expected percentage versus actual BTO percentage',20,23,'LineChart']];
  specs.forEach(x=>{let b=master.newChart()[x[5]==='PieChart'?'asPieChart':x[5]==='LineChart'?'asLineChart':'asColumnChart']().addRange(x[0]);if(x[1])b.addRange(x[1]);if(x[2].startsWith('Monthly BTO'))b.addRange(mm.getRange('K1:K200'));master.insertChart(b.setPosition(x[3],x[4],0,0).setOption('title',x[2]).build());});
}
