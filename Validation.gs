function refreshDropdowns() {
  const settings=mbSheet_('Settings'), rangeRule=a=>SpreadsheetApp.newDataValidation().requireValueInRange(settings.getRange(a),true).setAllowInvalid(false).build();
  const specs={Promos:{'Promo Type':'A2:A100','Result':'K2:K100'},BTO:{'BTO Method':'I2:I100','Result':'K2:K100'},NonPromos:{'Non-Promo Type':'J2:J100','Result':'K2:K100'}};
  Object.keys(specs).forEach(name=>{
    const s=mbSheet_(name);s.getRange(2,1,Math.max(1,s.getMaxRows()-1),s.getLastColumn()).clearDataValidations();
    Object.keys(specs[name]).forEach(header=>s.getRange(2,mbColumn_(s,header),MB.ROWS-1,1).setDataValidation(rangeRule(specs[name][header])));
  });
  refreshAccountValidations_();
  mbSheet_('EV Breakdown').getRange('B5').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(MB.PERIODS,true).setAllowInvalid(false).build());
}

function refreshAccountValidations_() {
  let source=SpreadsheetApp.getActive().getRangeByName('Account_ID_List');if(!source){refreshAccountIdList_();source=SpreadsheetApp.getActive().getRangeByName('Account_ID_List');}
  const rule=SpreadsheetApp.newDataValidation().requireValueInRange(source,true).setAllowInvalid(false).build();
  ['Promos','BTO','NonPromos'].forEach(name=>{const s=mbSheet_(name);s.getRange(2,mbColumn_(s,'Account'),MB.ROWS-1,1).setDataValidation(rule);});
  const accounts=mbSheet_('Accounts');if(accounts.getLastRow()>=5)accounts.getRange(5,mbColumn_(accounts,'Account',4),accounts.getMaxRows()-4,1).setDataValidation(rule);
  const funds=SpreadsheetApp.getActive().getSheetByName('Funds Management');if(funds){const rows=Math.min(10,Math.max(1,funds.getLastRow())),values=funds.getRange(1,1,rows,Math.max(1,funds.getLastColumn())).getDisplayValues();values.forEach((row,i)=>row.forEach((h,j)=>{if(['Account','Account ID'].includes(String(h).trim()))funds.getRange(i+2,j+1,Math.max(1,funds.getMaxRows()-i-1),1).clearDataValidations().setDataValidation(rule);}));}
}

function calculatedValidationCount_(){
  const specs={Promos:['Notes','Estimated EV','Cash Change','Bonus Change'],BTO:['Notes','EV Taken','Bookie Change','Betfair Change'],NonPromos:['Notes','Expected EV','EV %','Bookie Change','Betfair Change']};let count=0;
  Object.keys(specs).forEach(name=>{const s=mbSheet_(name);specs[name].forEach(h=>{count+=s.getRange(2,mbColumn_(s,h),s.getMaxRows()-1,1).getDataValidations().flat().filter(Boolean).length;});});
  const health=mbSheet_('Bookie Health');for(let p=0;p<MB.BOOKIE_HEALTH_PROFILES;p++)count+=health.getRange(3,3+p*2,MB.BOOKIE_HEALTH_ROWS,1).getDataValidations().flat().filter(Boolean).length;
  return count;
}
