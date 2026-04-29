/**
 * QueryBuilder — pure SQL construction functions.
 *
 * No Apps Script dependencies except Utilities.formatDate for Date quoting,
 * which is isolated behind an optional formatDateFn parameter for testability.
 */

/**
 * Quotes a value for use in a SQL statement.
 *
 * - null / empty string  → 'NULL'
 * - number              → numeric literal (unquoted)
 * - Date (duck-typed)   → quoted formatted string via formatDateFn
 * - everything else     → escaped + single-quoted string
 *
 * @param {*} value  The value to quote.
 * @param {Function=} formatDateFn  Optional date formatter (value) → string.
 *     Defaults to Utilities.formatDate with script timezone.
 * @return {string} SQL literal representation.
 */
function quoteSqlValue(value, formatDateFn) {
  if (value === null || value === '') {
    return 'NULL';
  }
  if (typeof value === 'number') {
    return String(value);
  }
  if (typeof value === 'object' && value !== null && typeof value.getTime === 'function') {
    if (typeof formatDateFn === 'function') {
      return "'" + formatDateFn(value) + "'";
    }
    return "'" + Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss') + "'";
  }
  var escaped = String(value).replace(/'/g, "''");
  return "'" + escaped + "'";
}

/**
 * Inverse of quoteSqlValue — parses a SQL literal back to a JS value.
 * Used for round-trip property testing.
 *
 * - 'NULL'                    → null
 * - Numeric string            → number
 * - Single-quoted string      → unescaped string ('' → ')
 *
 * @param {string} quoted  The SQL literal string.
 * @return {*} The parsed JavaScript value.
 */
function parseQuotedValue(quoted) {
  if (quoted === 'NULL') {
    return null;
  }
  // Check if it's a quoted string (starts and ends with single quote)
  if (quoted.length >= 2 && quoted.charAt(0) === "'" && quoted.charAt(quoted.length - 1) === "'") {
    var inner = quoted.substring(1, quoted.length - 1);
    return inner.replace(/''/g, "'");
  }
  // Try to parse as a number
  var num = Number(quoted);
  if (!isNaN(num) && quoted !== '') {
    return num;
  }
  // Fallback — return as-is
  return quoted;
}

/**
 * Builds a parameterized INSERT statement.
 *
 * @param {string} table     Table name.
 * @param {string[]} columns Column names.
 * @param {Array} values     Values corresponding to columns.
 * @return {{sql: string, params: Array}} Parameterized query.
 */
function buildInsertSql(table, columns, values) {
  var placeholders = [];
  for (var i = 0; i < columns.length; i++) {
    placeholders.push('?');
  }
  var sql = 'INSERT INTO ' + table + ' (' + columns.join(',') + ') VALUES (' + placeholders.join(',') + ')';
  return { sql: sql, params: values.slice(0, columns.length) };
}

/**
 * Builds a parameterized UPDATE statement.
 *
 * @param {string} table       Table name.
 * @param {string[]} columns   All column names.
 * @param {Array} values       All values corresponding to columns.
 * @param {number[]} setCols   Indices into columns/values for the SET clause.
 * @param {number[]} whereCols Indices into columns/values for the WHERE clause.
 * @return {{sql: string, params: Array}} Parameterized query.
 */
function buildUpdateSql(table, columns, values, setCols, whereCols) {
  var setParts = [];
  var params = [];
  for (var i = 0; i < setCols.length; i++) {
    setParts.push(columns[setCols[i]] + '=?');
    params.push(values[setCols[i]]);
  }
  var whereParts = [];
  for (var j = 0; j < whereCols.length; j++) {
    whereParts.push(columns[whereCols[j]] + '=?');
    params.push(values[whereCols[j]]);
  }
  var sql = 'UPDATE ' + table + ' SET ' + setParts.join(',') + ' WHERE ' + whereParts.join(' AND ');
  return { sql: sql, params: params };
}

/**
 * Builds a parameterized DELETE statement.
 *
 * @param {string} table       Table name.
 * @param {string[]} columns   All column names.
 * @param {Array} values       All values corresponding to columns.
 * @param {number[]} whereCols Indices into columns/values for the WHERE clause.
 * @return {{sql: string, params: Array}} Parameterized query.
 */
function buildDeleteSql(table, columns, values, whereCols) {
  var whereParts = [];
  var params = [];
  for (var i = 0; i < whereCols.length; i++) {
    whereParts.push(columns[whereCols[i]] + '=?');
    params.push(values[whereCols[i]]);
  }
  var sql = 'DELETE FROM ' + table + ' WHERE ' + whereParts.join(' AND ');
  return { sql: sql, params: params };
}
