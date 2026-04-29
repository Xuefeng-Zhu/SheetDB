var SQL_SHEET_NAME = 'SQL';
var CONFIG_MENU_NAME = 'Configure';

function ensureSheetByName(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (sheet == null) {
    sheet = ss.insertSheet(sheetName, ss.getNumSheets());
  }
  return sheet;
}

function ensureSqlSheet() {
  return ensureSheetByName(SQL_SHEET_NAME);
}

function quoteSqlValue(value) {
  if (value === null || value === '') {
    return 'NULL';
  }
  if (typeof value === 'number') {
    return String(value);
  }
  if (value instanceof Date) {
    return "'" + Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss") + "'";
  }
  var escaped = String(value).replace(/'/g, "''");
  return "'" + escaped + "'";
}
