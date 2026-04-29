var CONFIG_SHEET_NAME = 'Configue';

function getConfigSheet() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(CONFIG_SHEET_NAME);

  if (sheet == null) {
    sheet = ss.insertSheet(CONFIG_SHEET_NAME, 0);
    sheet.getRange(1, 1, 3, 1).setValues([
      ['URL'],
      ['Admin'],
      ['Password']
    ]);
    sheet.getRange(1, 1, 3, 1).setFontWeight('bold');
  }

  return sheet;
}

function getConfigValues() {
  var sheet = getConfigSheet();
  return {
    url: sheet.getRange(1, 2).getValue(),
    admin: sheet.getRange(2, 2).getValue(),
    password: sheet.getRange(3, 2).getValue()
  };
}

function hasConfigValues(config) {
  return config.url !== '' && config.admin !== '' && config.password !== '';
}
