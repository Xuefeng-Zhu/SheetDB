/**
 * ResultRenderer — writes query results to spreadsheet cells.
 *
 * renderReadResults, renderWriteResult, renderNoResults, replaceResultBlock
 */

/**
 * Writes a read-query result block to the sheet.
 * Appends a status row (sql + " Success"), then one row per data row,
 * followed by a blank separator row.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Target sheet
 * @param {string} sql - The SQL statement that was executed
 * @param {string[][]} rows - Array of row arrays (data rows from the query)
 */
function renderReadResults(sheet, sql, rows) {
  sheet.appendRow([sql + ' Success']);
  for (var i = 0; i < rows.length; i++) {
    sheet.appendRow(rows[i]);
  }
  sheet.appendRow([' ']);
}

/**
 * Writes a write-query result block to the sheet.
 * Appends a status row (sql + " Success") and a blank separator row.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Target sheet
 * @param {string} sql - The SQL statement that was executed
 */
function renderWriteResult(sheet, sql) {
  sheet.appendRow([sql + ' Success']);
  sheet.appendRow([' ']);
}

/**
 * Writes a status row indicating the query returned zero results.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Target sheet
 * @param {string} sql - The SQL statement that was executed
 */
function renderNoResults(sheet, sql) {
  sheet.appendRow([sql + ' — No results']);
}

/**
 * Replaces an existing result block in-place (used by Refresh).
 *
 * Starting from startRow, counts rows until hitting a blank row
 * (cell A value is ' ' or blank) or end of data. Deletes those old rows,
 * then inserts the new rows at the same position.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Target sheet
 * @param {number} startRow - 1-based row index where the result block begins
 * @param {any[][]} newRows - Array of row arrays to insert
 */
function replaceResultBlock(sheet, startRow, newRows) {
  // Count old block rows (same pattern as the original refresh logic)
  var nrow = 1;
  while (
    sheet.getRange(startRow + nrow, 1).getValue() !== ' ' &&
    !sheet.getRange(startRow + nrow, 1).isBlank()
  ) {
    nrow++;
  }

  // Adjust row count: insert or delete rows to match newRows length
  if (nrow < newRows.length) {
    sheet.insertRows(startRow, newRows.length - nrow);
  } else if (nrow > newRows.length) {
    sheet.deleteRows(startRow, nrow - newRows.length);
  }

  // Clear the status row and rewrite it
  sheet.deleteRow(startRow);
  sheet.insertRows(startRow);
  sheet.getRange(startRow, 1).setValue(newRows[0]);

  // Write remaining data rows using bulk setValues when possible
  var dataRows = newRows.slice(1);
  if (dataRows.length > 0) {
    var range = sheet.getRange(startRow + 1, 1, dataRows.length, dataRows[0].length);
    range.setValues(dataRows);
  }
}
