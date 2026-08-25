function onOpen() {
  SpreadsheetApp.getUi().createMenu('Matched Betting Tracker')
    .addItem('Build/Rebuild Workbook','buildWorkbook').addSeparator()
    .addItem('Refresh Formulas','refreshFormulas').addItem('Refresh Dropdowns','refreshDropdowns')
    .addItem('Refresh Dashboard','refreshDashboard').addItem('Run Audit','runAudit').addToUi();
}
