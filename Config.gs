const MB = Object.freeze({
  ROWS: 2000,
  SHEETS: ['Settings', 'Promos', 'BTO', 'NonPromos', 'Bookie Health', 'Accounts', 'Monthly Metrics', 'EV Breakdown', 'Master PnL', 'Audit'],
  BOOKIE_HEALTH_PROFILES: 20,
  BOOKIE_HEALTH_SEPARATOR: ', ',
  BOOKIE_HEALTH_STATUSES: [
    ['Open','#d9d2e9','#000000',1,true],
    ['Racing5','#d9d9d9','#000000',2,true],
    ['Racing10','#fce5cd','#000000',3,true],
    ['Racing25','#ffe599','#000000',4,true],
    ['Racing50','#b6d7a8','#000000',5,true],
    ['Sport5','#d9d9d9','#000000',6,true],
    ['Sport10','#f6b26b','#000000',7,true],
    ['Sport25','#ffd966','#000000',8,true],
    ['Sport50','#6aa84f','#000000',9,true],
    ['Unverified','#eeeeee','#000000',10,true],
    ['BAN','#000000','#ffffff',11,true],
    ['StatDec','#990000','#ffffff',12,true]
  ],
  PROMO_HEADERS: ['Entry ID','Date','Account','Promo Type','Notes','Turnover','Odds','Result','Promo EV %','Estimated EV','Cash Change','Bonus Change','Actual PnL','POT %','EV Tier','Month'],
  BTO_HEADERS: ['Date','Account','BTO Method','Notes','Turnover','Odds','BSP','Result','EV Taken','EV %','Bookie Change','Betfair Lay Stake','Lay Odds','Commission %','Betfair Change','Actual Return','Actual BTO %'],
  NON_HEADERS: ['Entry ID','Date','Account','Non-Promo Type','Notes','Turnover','Odds','BSP','Result','Expected EV','EV %','Bookie Change','Betfair Change','Actual PnL','POT %','Month'],
  MONTH_HEADERS: ['Month','Promo Turnover','Promo Estimated EV','Promo EV %','Promo Actual PnL','Promo POT %','BTO Turnover','BTO EV Taken','Expected BTO %','BTO Actual Return','Actual BTO %','Non-Promo Turnover','Non-Promo EV','Non-Promo EV %','Non-Promo Actual PnL','Non-Promo POT %','Overall EV','Overall Actual PnL','Cumulative EV','Cumulative PnL'],
  PROMOS: [['SGM3 Leg',.25],['SGM4 Leg',.225],['Free Hit Single',.5],['Free Hit Multi',.45],['Third Back',.15],['Second Back',.3],['Fourth+ Back',.25],['SRM',.09],['Megaboost',.325],['Boosts',.15],['Multi 3 Leg',.25],['Multi 4 Leg',.14],['Additional Winnings',.09],['Bonus Boost Winnings',.02],['Racing Multi',.15],['Preoutlay Third Back',.13],['Other',.15],['50% Back Third',.03],['Second = Win',.4]],
  BTO_METHODS: ['Horses','Greyhounds','Trots','First Score','Greeny','Singles','Correct Score'],
  NON_TYPES: ['Horses','Dogs','Trots','Two-way Dutch','Sport Single','Sport Multi','Fudge','Exotics','Betsniper EV','Betcloud TO'],
  RESULTS: ['W','L','B','Void'],
  PERIODS: ['Last 30 Days','Last 3 Months','Last 6 Months','Last 12 Months','All Time','Custom']
});

function mbSheet_(name) {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function mbEnsureSize_(sheet, rows, cols) {
  if (sheet.getMaxRows() < rows) sheet.insertRowsAfter(sheet.getMaxRows(), rows - sheet.getMaxRows());
  if (sheet.getMaxColumns() < cols) sheet.insertColumnsAfter(sheet.getMaxColumns(), cols - sheet.getMaxColumns());
}
