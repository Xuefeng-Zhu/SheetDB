/**
 * SqlPrompt — handles the SQL prompt dialog, refresh, and clear history.
 *
 * Depends on: InputValidator (validateSql, classifySql, isCancel),
 *             JdbcService (executeRead, executeWrite),
 *             ResultRenderer (renderReadResults, renderWriteResult, renderNoResults, replaceResultBlock),
 *             Helpers (appendSqlHistory, requireConfig)
 */

/**
 * Shows an input dialog for entering a SQL statement, validates, executes,
 * renders results on the active sheet, and records history.
 * If the user cancels, takes no action.
 */
function showPrompt() {
  requireConfig();

  var input = Browser.inputBox(
    'Google Sheet based SQL',
    'Please enter SQL statement you want to execute:',
    Browser.Buttons.OK_CANCEL);

  if (isCancel(input)) {
    return;
  }

  var sql = validateSql(input);
  var type = classifySql(sql);
  var sheet = SpreadsheetApp.getActiveSheet();

  try {
    if (type === 'read') {
      var rows = executeRead(sql);
      if (rows && rows.length > 0) {
        renderReadResults(sheet, sql, rows);
      } else {
        renderNoResults(sheet, sql);
      }
    } else {
      executeWrite(sql);
      renderWriteResult(sheet, sql);
    }
    appendSqlHistory([[sql + ' Success']]);
  } catch (e) {
    Browser.msgBox('SQL execution failed: ' + e.message);
  }
}

/**
 * Re-executes the SQL statement found in the currently selected cell
 * and replaces the previous result block in place.
 * Strips " Success" or " Failed" suffix from the cell value before re-executing.
 */
function refresh() {
  var cell = SpreadsheetApp.getActiveRange();
  var statement = cell.getValue();

  // Strip " Success" or " Failed" suffix
  statement = statement.replace(/ Success$/, '').replace(/ Failed$/, '').trim();

  var type = classifySql(statement);
  var sheet = cell.getSheet();
  var startRow = cell.getRowIndex();

  try {
    var newRows;
    if (type === 'read') {
      var rows = executeRead(statement);
      newRows = [statement + ' Success'];
      for (var i = 0; i < rows.length; i++) {
        newRows.push(rows[i]);
      }
    } else {
      executeWrite(statement);
      newRows = [statement + ' Success'];
    }

    replaceResultBlock(sheet, startRow, newRows);
  } catch (e) {
    Browser.msgBox('SQL execution failed: ' + e.message);
  }
}

/**
 * Confirms with the user, then clears all content from the active sheet.
 */
function clearHistory() {
  var result = Browser.msgBox(
    'Please confirm',
    'Are you sure you want to clear all the history?',
    Browser.Buttons.YES_NO);

  if (result === 'yes') {
    var sheet = SpreadsheetApp.getActiveSheet();
    sheet.clear();
    Browser.msgBox('History Cleared.');
  } else {
    Browser.msgBox('User Canceled.');
  }
}
