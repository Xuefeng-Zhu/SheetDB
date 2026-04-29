/**
 * InputValidator — pure validation functions.
 * No Apps Script dependencies; testable under Node.js.
 */

/**
 * Trims the input string and validates it is non-empty.
 * @param {string} input - Raw SQL input from the user.
 * @returns {string} Trimmed SQL string.
 * @throws {Error} If input is empty or whitespace-only.
 */
function validateSql(input) {
  if (typeof input !== 'string') {
    throw new Error('SQL input must be a string');
  }
  var trimmed = input.trim();
  if (trimmed.length === 0) {
    throw new Error('SQL statement cannot be empty or whitespace-only');
  }
  return trimmed;
}

/**
 * Classifies a SQL statement as 'read' or 'write' based on its first keyword.
 * @param {string} sql - A trimmed SQL string.
 * @returns {'read'|'write'} The classification.
 */
function classifySql(sql) {
  var firstKeyword = sql.trim().split(/\s+/)[0].toUpperCase();
  if (firstKeyword === 'SELECT' || firstKeyword === 'SHOW' || firstKeyword === 'DESCRIBE') {
    return 'read';
  }
  return 'write';
}

/**
 * Returns true if the value matches the cancel button sentinel.
 * @param {*} value - The value to check.
 * @returns {boolean}
 */
function isCancel(value) {
  return value === CANCEL_BUTTON_VALUE;
}
