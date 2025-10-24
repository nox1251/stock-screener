function initControlCenterSheet() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Control Center') || ss.insertSheet('Control Center');
  sh.clear();

  // Top config
  sh.getRange('A1:B1').setValues([['Setting','Value']]).setFontWeight('bold');
  sh.getRange('A2:B7').setValues([
    ['API Token',''],
    ['Active Exchange','PSE'],
    ['Retain Annual Years','10'],
    ['Retain Quarters','40'],
    ['Last Full Refresh',''],
    ['Last Extracted At','']   // <-- NEW
  ]);
  sh.setColumnWidths(1, 2, 220);

  // Main table (no Sheet Target)
  sh.getRange('A9:F9').setValues([[
    'Section', 'JSON Path', 'Field Name', 'Period Type', 'Active?', 'Notes'
  ]]).setFontWeight('bold');

  sh.setFrozenRows(9);
  sh.setFrozenColumns(1);
  sh.getRange('A10:F50').clearContent();

  try { applyControlCenterValidations(); } catch(e) {}

  SpreadsheetApp.getActive().toast(
    'Control Center initialized.\nNext: Refresh Data Dictionary → Apply Control Center Validations.',
    'Stock Screener',
    6
  );
}
