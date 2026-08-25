function refreshDropdowns() {
  const settings=mbSheet_('Settings');
  const rule = a => SpreadsheetApp.newDataValidation().requireValueInRange(settings.getRange(a),true).setAllowInvalid(false).build();
  mbSheet_('Promos').getRange(2,3,MB.ROWS-1,1).setDataValidation(rule('H2:H100'));
  mbSheet_('Promos').getRange(2,4,MB.ROWS-1,1).setDataValidation(rule('A2:A100'));
  mbSheet_('Promos').getRange(2,8,MB.ROWS-1,1).setDataValidation(rule('K2:K100'));
  mbSheet_('BTO').getRange(2,3,MB.ROWS-1,1).setDataValidation(rule('H2:H100'));
  mbSheet_('BTO').getRange(2,4,MB.ROWS-1,1).setDataValidation(rule('I2:I100'));
  mbSheet_('BTO').getRange(2,9,MB.ROWS-1,1).setDataValidation(rule('K2:K100'));
  mbSheet_('NonPromos').getRange(2,3,MB.ROWS-1,1).setDataValidation(rule('H2:H100'));
  mbSheet_('NonPromos').getRange(2,4,MB.ROWS-1,1).setDataValidation(rule('J2:J100'));
  mbSheet_('NonPromos').getRange(2,9,MB.ROWS-1,1).setDataValidation(rule('K2:K100'));
  mbSheet_('EV Breakdown').getRange('B2').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(MB.PERIODS,true).build());
}
