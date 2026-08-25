function buildCharts_() {
  const mm=mbSheet_('Monthly Metrics');[mm,mbSheet_('Master PnL')].forEach(s=>s.getCharts().forEach(c=>s.removeChart(c)));
  SpreadsheetApp.flush();const last=Math.max(2,mm.getRange('A:A').getDisplayValues().filter(r=>r[0]).length);
  mm.insertChart(mm.newChart().asLineChart().addRange(mm.getRange(1,1,last,1)).addRange(mm.getRange(1,19,last,2)).setNumHeaders(1).setPosition(last+3,2,0,0).setOption('title','Cumulative EV versus cumulative PnL').setOption('hAxis',{title:'Month',format:'MMM yyyy'}).setOption('vAxis',{title:'AUD',format:'currency'}).setOption('useFirstColumnAsDomain',true).build());
}
