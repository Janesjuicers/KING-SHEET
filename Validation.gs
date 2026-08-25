function refreshDropdowns() {
  const settings=mbSheet_('Settings');
  const rule = a => SpreadsheetApp.newDataValidation().requireValueInRange(settings.getRange(a),true).setAllowInvalid(false).build();
  refreshAccountValidations_();
  mbSheet_('Promos').getRange(2,3,MB.ROWS-1,1).setDataValidation(rule('A2:A100'));
  mbSheet_('Promos').getRange(2,7,MB.ROWS-1,1).setDataValidation(rule('K2:K100'));
  mbSheet_('BTO').getRange(2,3,MB.ROWS-1,1).setDataValidation(rule('I2:I100'));
  mbSheet_('BTO').getRange(2,8,MB.ROWS-1,1).setDataValidation(rule('K2:K100'));
  mbSheet_('NonPromos').getRange(2,3,MB.ROWS-1,1).setDataValidation(rule('J2:J100'));
  mbSheet_('NonPromos').getRange(2,8,MB.ROWS-1,1).setDataValidation(rule('K2:K100'));
  mbSheet_('EV Breakdown').getRange('B2').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(MB.PERIODS,true).build());
}

/** Reusable Account_ID_List validation for every account-entry column. */
function refreshAccountValidations_() {
  let source=SpreadsheetApp.getActive().getRangeByName('Account_ID_List');
  if(!source) { refreshAccountIdList_(); source=SpreadsheetApp.getActive().getRangeByName('Account_ID_List'); }
  const rule=SpreadsheetApp.newDataValidation().requireValueInRange(source,true).setAllowInvalid(false).build();
  [['Promos',2],['BTO',2],['NonPromos',2]].forEach(x=>mbSheet_(x[0]).getRange(2,x[1],MB.ROWS-1,1).setDataValidation(rule));
  const accounts=mbSheet_('Accounts'), headers=accounts.getRange(1,1,1,Math.max(1,accounts.getLastColumn())).getDisplayValues()[0];
  const col=Math.max(1,headers.indexOf('Account')+1); accounts.getRange(2,col,Math.max(MB.ROWS-1,accounts.getMaxRows()-1),1).setDataValidation(rule);
  const funds=SpreadsheetApp.getActive().getSheetByName('Funds Management');
  if(funds){const hs=funds.getRange(1,1,Math.min(10,Math.max(1,funds.getLastRow())),Math.max(1,funds.getLastColumn())).getDisplayValues();hs.forEach((row,i)=>row.forEach((h,j)=>{if(/^Account(?: ID)?$/i.test(String(h).trim()))funds.getRange(i+2,j+1,Math.max(MB.ROWS-1,funds.getMaxRows()-i-1),1).setDataValidation(rule);}));}
}
