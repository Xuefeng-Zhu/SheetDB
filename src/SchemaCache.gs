/**
 * SchemaCache — in-memory cache for table column metadata.
 * Valid for one script execution (Apps Script has no persistent in-memory state).
 * Depends on: JdbcService (executeRead).
 */

/** @type {Object.<string, string[]>} */
var schemaCache_ = {};

/**
 * Returns cached column names for the given table, or fetches via
 * DESCRIBE and caches the result.
 * @param {string} tableName - the database table name
 * @returns {string[]} array of column names
 */
function getColumns(tableName) {
  if (schemaCache_[tableName]) {
    return schemaCache_[tableName];
  }
  var rows = executeRead('DESCRIBE ' + tableName);
  var columns = [];
  for (var i = 0; i < rows.length; i++) {
    columns.push(rows[i][0]);
  }
  schemaCache_[tableName] = columns;
  return columns;
}

/**
 * Clears the schema cache for one table or all tables.
 * @param {string} [tableName] - if provided, clears only that table's cache entry;
 *                                otherwise clears the entire cache.
 */
function invalidate(tableName) {
  if (tableName !== undefined && tableName !== null) {
    delete schemaCache_[tableName];
  } else {
    schemaCache_ = {};
  }
}
