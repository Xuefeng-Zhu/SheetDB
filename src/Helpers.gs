/**
 * Helpers — shared utility functions that depend on Apps Script APIs.
 *
 * ensureSheet, ensureSqlSheet, appendSqlHistory, requireConfig
 */

/**
 * Returns the sheet with the given name, creating it if it does not exist.
 * @param {string} name - Sheet name
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function ensureSheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (sheet === null) {
    sheet = ss.insertSheet(name, ss.getNumSheets());
  }
  return sheet;
}

/**
 * Shorthand for ensureSheet(SHEET_NAME_SQL).
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function ensureSqlSheet() {
  return ensureSheet(SHEET_NAME_SQL);
}

/**
 * Appends rows to the SQL history sheet.
 * Returns early if rows is null, undefined, or empty.
 * @param {any[][]} rows - Array of row arrays to append
 */
function appendSqlHistory(rows) {
  if (!rows || rows.length === 0) return;
  var sqlSheet = ensureSqlSheet();
  for (var i = 0; i < rows.length; i++) {
    sqlSheet.appendRow(rows[i]);
  }
}

/**
 * Gets config via getConfig(), validates with hasValidConfig().
 * If not configured, shows MSG_CONFIGURE_FIRST error dialog and throws.
 * Otherwise returns the config object.
 * @returns {{url: string, user: string, password: string}}
 */
function requireConfig() {
  var config = getConfig();
  if (!hasValidConfig(config)) {
    Browser.msgBox(MSG_CONFIGURE_FIRST);
    throw new Error(MSG_CONFIGURE_FIRST);
  }
  return config;
}
