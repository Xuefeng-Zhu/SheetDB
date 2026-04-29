/**
 * JdbcService — unified JDBC execution functions.
 * Single module for all JDBC operations.
 * Depends on: Jdbc (Apps Script), ConfigManager (getConfig), Constants (JDBC_TIMEOUT_SECONDS).
 *
 * All public functions throw errors with the original message and the failing SQL.
 * Close errors in finally blocks are suppressed (logged to console.error).
 */

/**
 * Internal. Gets config via getConfig(), opens a JDBC connection,
 * and sets the statement timeout to JDBC_TIMEOUT_SECONDS.
 * @returns {JdbcConnection}
 * @throws {Error} with context on connection failure
 */
function getConnection_() {
  var config = getConfig();
  try {
    var conn = Jdbc.getConnection(config.url, config.user, config.password);
    return conn;
  } catch (e) {
    throw new Error('JDBC connection failed: ' + e.message);
  }
}

/**
 * Opens connection, executes a read query, returns array of row arrays.
 * Closes ResultSet, Statement, and Connection in finally block.
 * @param {string} sql - SQL query to execute (SELECT, SHOW, DESCRIBE)
 * @returns {string[][]} array of row arrays
 * @throws {Error} with original message + failing SQL
 */
function executeRead(sql) {
  var conn = null, stmt = null, rs = null;
  try {
    conn = getConnection_();
    stmt = conn.createStatement();
    stmt.setQueryTimeout(JDBC_TIMEOUT_SECONDS);
    rs = stmt.executeQuery(sql);
    var meta = rs.getMetaData();
    var numCols = meta.getColumnCount();
    var results = [];
    while (rs.next()) {
      var row = [];
      for (var col = 1; col <= numCols; col++) {
        row.push(rs.getString(col));
      }
      results.push(row);
    }
    return results;
  } catch (e) {
    throw new Error('SQL failed [' + sql + ']: ' + e.message);
  } finally {
    try { if (rs) rs.close(); } catch (e) { console.error('close rs: ' + e.message); }
    try { if (stmt) stmt.close(); } catch (e) { console.error('close stmt: ' + e.message); }
    try { if (conn) conn.close(); } catch (e) { console.error('close conn: ' + e.message); }
  }
}

/**
 * Opens connection, executes a write statement, returns affected row count.
 * Closes Statement and Connection in finally block.
 * @param {string} sql - SQL statement to execute (INSERT, UPDATE, DELETE, etc.)
 * @returns {number} affected row count
 * @throws {Error} with original message + failing SQL
 */
function executeWrite(sql) {
  var conn = null, stmt = null;
  try {
    conn = getConnection_();
    stmt = conn.createStatement();
    stmt.setQueryTimeout(JDBC_TIMEOUT_SECONDS);
    var count = stmt.executeUpdate(sql);
    return count;
  } catch (e) {
    throw new Error('SQL failed [' + sql + ']: ' + e.message);
  } finally {
    try { if (stmt) stmt.close(); } catch (e) { console.error('close stmt: ' + e.message); }
    try { if (conn) conn.close(); } catch (e) { console.error('close conn: ' + e.message); }
  }
}

/**
 * Opens connection, creates PreparedStatement, binds params by type,
 * executes update, returns affected row count.
 * Closes PreparedStatement and Connection in finally block.
 * @param {string} sql - SQL with ? placeholders
 * @param {any[]} params - parameter values to bind
 * @returns {number} affected row count
 * @throws {Error} with original message + failing SQL
 */
function executePrepared(sql, params) {
  var conn = null, pstmt = null;
  try {
    conn = getConnection_();
    pstmt = conn.prepareStatement(sql);
    pstmt.setQueryTimeout(JDBC_TIMEOUT_SECONDS);
    for (var i = 0; i < params.length; i++) {
      var value = params[i];
      if (value === null || value === undefined) {
        pstmt.setNull(i + 1, 0);
      } else if (typeof value === 'number') {
        if (Number.isInteger(value)) {
          pstmt.setInt(i + 1, value);
        } else {
          pstmt.setDouble(i + 1, value);
        }
      } else if (value instanceof Date) {
        var formatted = Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
        pstmt.setString(i + 1, formatted);
      } else {
        pstmt.setString(i + 1, String(value));
      }
    }
    var count = pstmt.executeUpdate();
    return count;
  } catch (e) {
    throw new Error('SQL failed [' + sql + ']: ' + e.message);
  } finally {
    try { if (pstmt) pstmt.close(); } catch (e) { console.error('close pstmt: ' + e.message); }
    try { if (conn) conn.close(); } catch (e) { console.error('close conn: ' + e.message); }
  }
}
