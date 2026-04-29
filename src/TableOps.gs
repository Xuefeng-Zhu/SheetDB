/**
 * TableOps — handles table loading, refresh, drop, and bulk operations.
 *
 * Depends on: JdbcService (executeRead, executeWrite, executePrepared),
 *             SchemaCache (getColumns, invalidate),
 *             Helpers (appendSqlHistory, requireConfig, ensureSheet),
 *             InputValidator (isCancel),
 *             Constants (SHEET_NAME_SQL)
 */

/**
 * Fetches the list of tables from the database and creates or refreshes
 * a Table_Sheet for each table. Column headers are bold in row 1,
 * data rows start from row 2.
 */
function loadTables() {
  requireConfig();

  var tables = executeRead('SHOW tables');
  if (!tables || tables.length === 0) {
    return;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  for (var i = 0; i < tables.length; i++) {
    var tableName = tables[i][0];
    var sheet = ss.getSheetByName(tableName);

    if (sheet !== null) {
      // Sheet exists — clear and reload
      sheet.clear();
    } else {
      // Create new sheet
      sheet = ss.insertSheet(tableName, ss.getNumSheets());
    }

    // Write bold headers in row 1
    var columns = getColumns(tableName);
    sheet.appendRow(columns);
    var headerRange = sheet.getRange(1, 1, 1, columns.length);
    headerRange.setFontWeight('bold');

    // Fetch and write data starting from row 2
    var data = executeRead('SELECT * FROM ' + tableName);
    if (data && data.length > 0) {
      var dataRange = sheet.getRange(2, 1, data.length, data[0].length);
      dataRange.setValues(data);
    }
  }
}

/**
 * Reloads schema and data for the current or prompted table.
 * If on the SQL_Sheet, prompts for a table name; otherwise uses the
 * current sheet name. Clears the sheet and reloads headers + data.
 * Invalidates the schema cache for the table.
 */
function refreshTable() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var tableName;

  if (sheet.getSheetName() === SHEET_NAME_SQL) {
    var result = Browser.inputBox(
      'Refresh Table',
      'Please enter table name for refresh:',
      Browser.Buttons.OK_CANCEL);

    if (isCancel(result)) {
      Browser.msgBox('Thanks for using! Bye!');
      return;
    }

    tableName = result;
    sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tableName);
  } else {
    tableName = sheet.getSheetName();
  }

  if (!sheet) {
    Browser.msgBox('Table not found.');
    return;
  }

  try {
    // Invalidate schema cache for this table
    invalidate(tableName);

    // Clear and reload
    sheet.clear();

    var columns = getColumns(tableName);
    sheet.appendRow(columns);
    var headerRange = sheet.getRange(1, 1, 1, columns.length);
    headerRange.setFontWeight('bold');

    var data = executeRead('SELECT * FROM ' + tableName);
    if (data && data.length > 0) {
      var dataRange = sheet.getRange(2, 1, data.length, data[0].length);
      dataRange.setValues(data);
    }

    Browser.msgBox('Table successfully refreshed');
  } catch (e) {
    Browser.msgBox('Refresh failed: ' + e.message);
  }
}

/**
 * Drops a database table and removes its sheet.
 * If on the SQL_Sheet, prompts for a table name.
 * Records the DROP statement in SQL history.
 * Handles cancel and errors gracefully.
 */
function dropTable() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var tableName;

  if (sheet.getSheetName() === SHEET_NAME_SQL) {
    var result = Browser.inputBox(
      'Drop Table',
      'Please enter table name for drop:',
      Browser.Buttons.OK_CANCEL);

    if (isCancel(result)) {
      Browser.msgBox('Thanks for using! Bye!');
      return;
    }

    tableName = result;
    sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tableName);
  } else {
    tableName = sheet.getSheetName();
  }

  if (!sheet) {
    Browser.msgBox('Table not found.');
    return;
  }

  try {
    var sql = 'DROP TABLE ' + tableName;
    executeWrite(sql);
    appendSqlHistory([[sql + ' Success']]);
    SpreadsheetApp.getActiveSpreadsheet().deleteSheet(sheet);
  } catch (e) {
    Browser.msgBox('Drop table failed: ' + e.message);
  }
}

// ---------------------------------------------------------------------------
// Bulk Report helpers (pure functions — testable under Node.js)
// ---------------------------------------------------------------------------

/**
 * Creates a new BulkReport object for tracking bulk operation outcomes.
 * @param {number} total - Total number of rows to process.
 * @returns {{total: number, succeeded: number, failed: number, errors: string[]}}
 */
function createBulkReport(total) {
  return { total: total, succeeded: 0, failed: 0, errors: [] };
}

/**
 * Records a single row outcome in the BulkReport.
 * @param {{total: number, succeeded: number, failed: number, errors: string[]}} report
 * @param {boolean} success - Whether the row operation succeeded.
 * @param {string=} errorMsg - Error message if the operation failed.
 */
function recordBulkResult(report, success, errorMsg) {
  if (success) {
    report.succeeded++;
  } else {
    report.failed++;
    if (errorMsg) {
      report.errors.push(errorMsg);
    }
  }
}

// ---------------------------------------------------------------------------
// Bulk operations: fastInsert, fastUpdate, fastDelete
// ---------------------------------------------------------------------------

/**
 * Bulk-inserts selected rows into a database table using parameterized queries.
 *
 * If on the SQL_Sheet, prompts for table name (and optional column list).
 * If on a Table_Sheet, uses the sheet name as the table and row-1 headers
 * as columns.
 *
 * Reads the currently selected range, builds one INSERT per row via
 * buildInsertSql, executes via executePrepared, and records each successful
 * INSERT in SQL history. Uses continue-on-error with BulkReport.
 */
function fastInsert() {
  requireConfig();

  var sheet = SpreadsheetApp.getActiveSheet();
  var tableName;
  var columns;

  if (sheet.getSheetName() === SHEET_NAME_SQL) {
    var result = Browser.inputBox(
      'Fast Insert',
      'Please enter table name (and attributes of inserted data separate by comma):',
      Browser.Buttons.OK_CANCEL);

    if (isCancel(result)) {
      Browser.msgBox('Thanks for using! Bye!');
      return;
    }

    var parts = result.split(',');
    tableName = parts[0].trim();
    if (parts.length > 1) {
      columns = [];
      for (var i = 1; i < parts.length; i++) {
        columns.push(parts[i].trim());
      }
    } else {
      columns = null;
    }
  } else {
    tableName = sheet.getSheetName();
    var headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    columns = headerRange.getValues()[0];
  }

  var range = SpreadsheetApp.getActiveRange();
  var values = range.getValues();
  var report = createBulkReport(values.length);

  for (var r = 0; r < values.length; r++) {
    try {
      var row = values[r];
      var rowCols = columns;

      // If no columns specified (SQL_Sheet without column list), build without column names
      if (!rowCols) {
        // Build a simple column-less insert: INSERT INTO table VALUES (?,?...)
        var placeholders = [];
        for (var p = 0; p < row.length; p++) {
          placeholders.push('?');
        }
        var sql = 'INSERT INTO ' + tableName + ' VALUES (' + placeholders.join(',') + ')';
        executePrepared(sql, row);
        appendSqlHistory([[sql + ' Success']]);
      } else {
        var query = buildInsertSql(tableName, rowCols, row);
        executePrepared(query.sql, query.params);
        appendSqlHistory([[query.sql + ' Success']]);
      }
      recordBulkResult(report, true);
    } catch (e) {
      recordBulkResult(report, false, 'Row ' + (r + 1) + ': ' + e.message);
    }
  }

  Browser.msgBox('Completed: ' + report.succeeded + ' succeeded, ' + report.failed + ' failed.');
}

/**
 * Bulk-updates selected rows in a database table using parameterized queries.
 *
 * If on the SQL_Sheet, prompts for table name and column list.
 * If on a Table_Sheet, uses the sheet name as the table and row-1 headers
 * as columns.
 *
 * Shows a checkbox dialog so the user can select which columns to SET
 * (the rest become WHERE conditions). Builds one UPDATE per row via
 * buildUpdateSql, executes via executePrepared. Uses continue-on-error
 * with BulkReport.
 */
function fastUpdate() {
  requireConfig();

  var sheet = SpreadsheetApp.getActiveSheet();
  var tableName;
  var columns;

  if (sheet.getSheetName() === SHEET_NAME_SQL) {
    var result = Browser.inputBox(
      'Fast Update',
      'Please enter table name and table attributes separate by comma:',
      Browser.Buttons.OK_CANCEL);

    if (isCancel(result)) {
      Browser.msgBox('Thanks for using! Bye!');
      return;
    }

    var parts = result.split(',');
    tableName = parts[0].trim();
    columns = [];
    for (var i = 1; i < parts.length; i++) {
      columns.push(parts[i].trim());
    }
  } else {
    tableName = sheet.getSheetName();
    var headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    columns = headerRange.getValues()[0];
  }

  // Build checkbox dialog HTML for column selection (SET columns)
  var htmlContent = '<html><head><style>'
    + 'body { font-family: Arial, sans-serif; padding: 10px; }'
    + 'label { display: block; margin: 5px 0; }'
    + 'button { margin-top: 10px; padding: 5px 15px; }'
    + '</style></head><body>'
    + '<p>Select columns to UPDATE (SET). Unselected columns will be used as WHERE conditions.</p>';

  for (var c = 0; c < columns.length; c++) {
    htmlContent += '<label><input type="checkbox" value="' + c + '"> ' + columns[c] + '</label>';
  }

  htmlContent += '<br><button onclick="submitSelection()">OK</button>'
    + '<script>'
    + 'function submitSelection() {'
    + '  var checkboxes = document.querySelectorAll("input[type=checkbox]:checked");'
    + '  var indices = [];'
    + '  for (var i = 0; i < checkboxes.length; i++) {'
    + '    indices.push(checkboxes[i].value);'
    + '  }'
    + '  google.script.host.close();'
    + '  google.script.run.fastUpdateWithSelection(indices.join(","));'
    + '}'
    + '</script></body></html>';

  var htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(300)
    .setHeight(400)
    .setTitle('Select SET Columns');

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Select SET Columns');
}

/**
 * Callback invoked by the fastUpdate checkbox dialog.
 * Receives a comma-separated string of column indices selected for SET.
 * Reads the active range and builds parameterized UPDATE statements.
 *
 * @param {string} selectionStr - Comma-separated column indices for SET clause.
 */
function fastUpdateWithSelection(selectionStr) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var tableName;
  var columns;

  if (sheet.getSheetName() === SHEET_NAME_SQL) {
    // In the SQL_Sheet context, we need to re-prompt or use a cached value.
    // Since the dialog callback is a separate invocation, we re-read from
    // a script property if needed. For simplicity, prompt again.
    var result = Browser.inputBox(
      'Fast Update',
      'Please enter table name and table attributes separate by comma:',
      Browser.Buttons.OK_CANCEL);

    if (isCancel(result)) {
      Browser.msgBox('Thanks for using! Bye!');
      return;
    }

    var parts = result.split(',');
    tableName = parts[0].trim();
    columns = [];
    for (var i = 1; i < parts.length; i++) {
      columns.push(parts[i].trim());
    }
  } else {
    tableName = sheet.getSheetName();
    var headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    columns = headerRange.getValues()[0];
  }

  // Parse selected SET column indices
  var setCols = [];
  if (selectionStr && selectionStr.length > 0) {
    var selectedParts = selectionStr.split(',');
    for (var s = 0; s < selectedParts.length; s++) {
      setCols.push(parseInt(selectedParts[s], 10));
    }
  }

  // Build WHERE column indices (all columns not in SET)
  var setColSet = {};
  for (var sc = 0; sc < setCols.length; sc++) {
    setColSet[setCols[sc]] = true;
  }
  var whereCols = [];
  for (var w = 0; w < columns.length; w++) {
    if (!setColSet[w]) {
      whereCols.push(w);
    }
  }

  if (setCols.length === 0) {
    Browser.msgBox('No columns selected for update.');
    return;
  }

  var range = SpreadsheetApp.getActiveRange();
  var values = range.getValues();
  var report = createBulkReport(values.length);

  for (var r = 0; r < values.length; r++) {
    try {
      var row = values[r];
      var query = buildUpdateSql(tableName, columns, row, setCols, whereCols);
      executePrepared(query.sql, query.params);
      appendSqlHistory([[query.sql + ' Success']]);
      recordBulkResult(report, true);
    } catch (e) {
      recordBulkResult(report, false, 'Row ' + (r + 1) + ': ' + e.message);
    }
  }

  Browser.msgBox('Completed: ' + report.succeeded + ' succeeded, ' + report.failed + ' failed.');
}

/**
 * Bulk-deletes selected rows from a database table using parameterized queries.
 *
 * If on the SQL_Sheet, prompts for table name and column list.
 * If on a Table_Sheet, uses the sheet name as the table and row-1 headers
 * as columns.
 *
 * Builds one DELETE per row via buildDeleteSql using all columns as WHERE
 * conditions, executes via executePrepared. On success, removes the
 * corresponding rows from the sheet. Uses continue-on-error with BulkReport.
 */
function fastDelete() {
  requireConfig();

  var sheet = SpreadsheetApp.getActiveSheet();
  var tableName;
  var columns;

  if (sheet.getSheetName() === SHEET_NAME_SQL) {
    var result = Browser.inputBox(
      'Fast Delete',
      'Please enter table name and table attributes separate by comma:',
      Browser.Buttons.OK_CANCEL);

    if (isCancel(result)) {
      Browser.msgBox('Thanks for using! Bye!');
      return;
    }

    var parts = result.split(',');
    tableName = parts[0].trim();
    columns = [];
    for (var i = 1; i < parts.length; i++) {
      columns.push(parts[i].trim());
    }
  } else {
    tableName = sheet.getSheetName();
    var headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    columns = headerRange.getValues()[0];
  }

  // Build WHERE column indices — use all columns
  var whereCols = [];
  for (var c = 0; c < columns.length; c++) {
    whereCols.push(c);
  }

  var range = SpreadsheetApp.getActiveRange();
  var values = range.getValues();
  var report = createBulkReport(values.length);

  // Track which rows succeeded for sheet row removal (process in reverse)
  var succeededRows = [];

  for (var r = 0; r < values.length; r++) {
    try {
      var row = values[r];
      var query = buildDeleteSql(tableName, columns, row, whereCols);
      executePrepared(query.sql, query.params);
      appendSqlHistory([[query.sql + ' Success']]);
      recordBulkResult(report, true);
      succeededRows.push(r);
    } catch (e) {
      recordBulkResult(report, false, 'Row ' + (r + 1) + ': ' + e.message);
    }
  }

  // Remove successfully deleted rows from the sheet (in reverse order to preserve indices)
  if (sheet.getSheetName() !== SHEET_NAME_SQL && succeededRows.length > 0) {
    var startRow = range.getRowIndex();
    for (var d = succeededRows.length - 1; d >= 0; d--) {
      sheet.deleteRow(startRow + succeededRows[d]);
    }
  }

  Browser.msgBox('Completed: ' + report.succeeded + ' succeeded, ' + report.failed + ' failed.');
}
