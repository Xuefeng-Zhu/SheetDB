var SQL_SHEET_NAME = 'SQL';
var CONFIG_MENU_NAME = 'Configure';
var CANCEL_BUTTON_VALUE = 'cancel';

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

function isCancel(inputValue) {
  return inputValue === CANCEL_BUTTON_VALUE;
}

function appendSqlHistory(rows) {
  if (!rows || rows.length === 0) return;
  var sqlSheet = ensureSqlSheet();
  for (var i = 0; i < rows.length; i++) {
    sqlSheet.appendRow(rows[i]);
  }
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
