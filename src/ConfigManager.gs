/**
 * ConfigManager — configuration management functions.
 * Manages the Config_Sheet for storing database credentials.
 *
 * Pure function: hasValidConfig (testable under Node.js)
 * Apps Script dependent: getConfigSheet, getConfig, saveConfig, runConfigDialog
 */

/**
 * Returns the Config_Sheet, creating it with labeled rows if absent.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getConfigSheet() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(SHEET_NAME_CONFIG);

  if (sheet === null) {
    sheet = ss.insertSheet(SHEET_NAME_CONFIG, 0);
    sheet.getRange(1, 1, 3, 1).setValues([
      ['URL'],
      ['Admin'],
      ['Password']
    ]);
    sheet.getRange(1, 1, 3, 1).setFontWeight('bold');
  }

  return sheet;
}

/**
 * Reads credentials from Config_Sheet.
 * @returns {{url: string, user: string, password: string}}
 */
function getConfig() {
  var sheet = getConfigSheet();
  return {
    url: sheet.getRange(1, 2).getValue(),
    user: sheet.getRange(2, 2).getValue(),
    password: sheet.getRange(3, 2).getValue()
  };
}

/**
 * Returns true if all three credential fields are non-empty strings.
 * Pure function — no Apps Script dependencies.
 * @param {{url: *, user: *, password: *}} config
 * @returns {boolean}
 */
function hasValidConfig(config) {
  return typeof config.url === 'string' && config.url !== '' &&
         typeof config.user === 'string' && config.user !== '' &&
         typeof config.password === 'string' && config.password !== '';
}

/**
 * Writes credentials to Config_Sheet.
 * @param {string} url - JDBC URL
 * @param {string} user - Database username
 * @param {string} password - Database password
 */
function saveConfig(url, user, password) {
  var sheet = getConfigSheet();
  sheet.getRange(1, 2).setValue(url);
  sheet.getRange(2, 2).setValue(user);
  sheet.getRange(3, 2).setValue(password);
}

/**
 * Prompts user for URL, username, password in sequence via Browser.inputBox.
 * Returns false if cancelled at any step, displays cancellation message.
 * On success, calls saveConfig.
 * @returns {boolean}
 */
function runConfigDialog() {
  var url = Browser.inputBox('Enter database URL:');
  if (isCancel(url)) {
    Browser.msgBox('Configuration cancelled.');
    return false;
  }

  var user = Browser.inputBox('Enter username:');
  if (isCancel(user)) {
    Browser.msgBox('Configuration cancelled.');
    return false;
  }

  var password = Browser.inputBox('Enter password:');
  if (isCancel(password)) {
    Browser.msgBox('Configuration cancelled.');
    return false;
  }

  saveConfig(url, user, password);
  return true;
}
