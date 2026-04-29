const fs = require('fs');
const vm = require('vm');
const fc = require('fast-check');

function loadQueryBuilder() {
  var constantsCode = fs.readFileSync('src/Constants.gs', 'utf8');
  var qbCode = fs.readFileSync('src/QueryBuilder.gs', 'utf8');
  var sandbox = {
    Utilities: {
      formatDate: function () { return '2026-01-02 03:04:05'; },
    },
    Session: {
      getScriptTimeZone: function () { return 'UTC'; },
    },
  };
  vm.createContext(sandbox);
  vm.runInContext(constantsCode, sandbox);
  vm.runInContext(qbCode, sandbox);
  return sandbox;
}

function loadConfigManager() {
  var constantsCode = fs.readFileSync('src/Constants.gs', 'utf8');
  var cmCode = fs.readFileSync('src/ConfigManager.gs', 'utf8');
  var sandbox = {
    SpreadsheetApp: { getActive: function () { return null; } },
    Browser: { inputBox: function () { return ''; }, msgBox: function () {} },
  };
  vm.createContext(sandbox);
  vm.runInContext(constantsCode, sandbox);
  vm.runInContext(cmCode, sandbox);
  return sandbox;
}

function loadInputValidator() {
  const constantsCode = fs.readFileSync('src/Constants.gs', 'utf8');
  const validatorCode = fs.readFileSync('src/InputValidator.gs', 'utf8');
  const sandbox = {};
  vm.createContext(sandbox);
  vm.runInContext(constantsCode, sandbox);
  vm.runInContext(validatorCode, sandbox);
  return sandbox;
}

function loadTableOps() {
  var constantsCode = fs.readFileSync('src/Constants.gs', 'utf8');
  var tableOpsCode = fs.readFileSync('src/TableOps.gs', 'utf8');
  var noop = function () {};
  var sandbox = {
    SpreadsheetApp: { getActiveSpreadsheet: noop, getActiveSheet: noop, getActiveRange: noop, getUi: noop },
    Browser: { inputBox: noop, msgBox: noop, Buttons: { OK_CANCEL: 0 } },
    HtmlService: { createHtmlOutput: function () { return { setWidth: function () { return { setHeight: function () { return { setTitle: function () { return {}; } }; } }; } }; } },
    requireConfig: noop,
    executeRead: noop,
    executeWrite: noop,
    executePrepared: noop,
    buildInsertSql: noop,
    buildUpdateSql: noop,
    buildDeleteSql: noop,
    appendSqlHistory: noop,
    isCancel: function () { return false; },
    getColumns: noop,
    invalidate: noop,
  };
  vm.createContext(sandbox);
  vm.runInContext(constantsCode, sandbox);
  vm.runInContext(tableOpsCode, sandbox);
  return sandbox;
}

function assertEqual(actual, expected, message) {
  if (actual !== expected) {
    throw new Error(`${message}\nExpected: ${expected}\nActual:   ${actual}`);
  }
}

(function run() {
  // --- Property-based tests ---
  const iv = loadInputValidator();

  /**
   * Feature: sheetdb-rebuild, Property 2: Whitespace SQL rejection
   * Validates: Requirements 3.1
   *
   * For any string composed entirely of whitespace characters,
   * validateSql(input) SHALL throw a descriptive error and never return a value.
   */
  fc.assert(
    fc.property(
      fc.array(fc.constantFrom(' ', '\t', '\n', '\r'), { minLength: 0, maxLength: 200 }).map(function (arr) { return arr.join(''); }),
      function (whitespaceStr) {
        var threw = false;
        try {
          iv.validateSql(whitespaceStr);
        } catch (e) {
          threw = true;
          if (typeof e.message !== 'string' || e.message.length === 0) {
            throw new Error('Expected a descriptive error message but got: ' + e.message);
          }
        }
        if (!threw) {
          throw new Error('validateSql should have thrown for whitespace-only input: ' + JSON.stringify(whitespaceStr));
        }
      }
    ),
    { numRuns: 100 }
  );
  console.log('Property 2 (Whitespace SQL rejection) passed — 100 iterations.');

  /**
   * Feature: sheetdb-rebuild, Property 3: SQL classification correctness
   * Validates: Requirements 14.4
   *
   * For any SQL string whose first non-whitespace keyword is SELECT, SHOW, or
   * DESCRIBE (case-insensitive), classifySql SHALL return 'read'. For any SQL
   * string whose first keyword is anything else (INSERT, UPDATE, DELETE, DROP,
   * CREATE, ALTER, etc.), classifySql SHALL return 'write'.
   */
  var readKeywords = ['SELECT', 'SHOW', 'DESCRIBE'];
  var writeKeywords = ['INSERT', 'UPDATE', 'DELETE', 'DROP', 'CREATE', 'ALTER'];

  var randomCaseKeyword = function (keyword) {
    return fc.array(fc.boolean(), { minLength: keyword.length, maxLength: keyword.length }).map(function (flags) {
      return keyword.split('').map(function (ch, i) {
        return flags[i] ? ch.toUpperCase() : ch.toLowerCase();
      }).join('');
    });
  };

  var sqlSuffixArb = fc.array(fc.constantFrom(
    ' ', '*', 'a', 'b', 'c', '1', '2', '_', ',', '(', ')', 'FROM', 'INTO', 'SET', 'WHERE', 'TABLE'
  ), { minLength: 0, maxLength: 30 }).map(function (arr) { return ' ' + arr.join(''); });

  var leadingWhitespaceArb = fc.array(fc.constantFrom(' ', '\t', '\n', '\r'), { minLength: 0, maxLength: 10 }).map(function (arr) { return arr.join(''); });

  var readSqlArb = fc.tuple(
    fc.constantFrom.apply(fc, readKeywords),
    leadingWhitespaceArb,
    sqlSuffixArb
  ).chain(function (tuple) {
    var keyword = tuple[0];
    var ws = tuple[1];
    var suffix = tuple[2];
    return randomCaseKeyword(keyword).map(function (randomized) {
      return { sql: ws + randomized + suffix, expected: 'read' };
    });
  });

  var writeSqlArb = fc.tuple(
    fc.constantFrom.apply(fc, writeKeywords),
    leadingWhitespaceArb,
    sqlSuffixArb
  ).chain(function (tuple) {
    var keyword = tuple[0];
    var ws = tuple[1];
    var suffix = tuple[2];
    return randomCaseKeyword(keyword).map(function (randomized) {
      return { sql: ws + randomized + suffix, expected: 'write' };
    });
  });

  fc.assert(
    fc.property(
      fc.oneof(readSqlArb, writeSqlArb),
      function (testCase) {
        var result = iv.classifySql(testCase.sql);
        if (result !== testCase.expected) {
          throw new Error(
            'classifySql(' + JSON.stringify(testCase.sql) + ') returned ' +
            JSON.stringify(result) + ' but expected ' + JSON.stringify(testCase.expected)
          );
        }
      }
    ),
    { numRuns: 100 }
  );
  console.log('Property 3 (SQL classification correctness) passed — 100 iterations.');

  // --- Example-based unit tests for InputValidator ---

  // validateSql: empty string should throw
  var threw = false;
  try { iv.validateSql(''); } catch (e) { threw = true; }
  assertEqual(threw, true, 'validateSql("") should throw');

  // validateSql: whitespace-only should throw
  threw = false;
  try { iv.validateSql('   '); } catch (e) { threw = true; }
  assertEqual(threw, true, 'validateSql("   ") should throw');

  // validateSql: tab/newline whitespace should throw
  threw = false;
  try { iv.validateSql('\t\n'); } catch (e) { threw = true; }
  assertEqual(threw, true, 'validateSql("\\t\\n") should throw');

  // validateSql: valid SQL with surrounding whitespace returns trimmed
  assertEqual(iv.validateSql('  SELECT * FROM t  '), 'SELECT * FROM t', 'validateSql should trim whitespace');

  // validateSql: valid SQL without extra whitespace returns as-is
  assertEqual(iv.validateSql('INSERT INTO t VALUES (1)'), 'INSERT INTO t VALUES (1)', 'validateSql should return valid SQL unchanged');

  // classifySql: SELECT -> read
  assertEqual(iv.classifySql('SELECT * FROM t'), 'read', 'classifySql SELECT should return read');

  // classifySql: lowercase select -> read (case insensitive)
  assertEqual(iv.classifySql('select * from t'), 'read', 'classifySql select (lowercase) should return read');

  // classifySql: SHOW -> read
  assertEqual(iv.classifySql('SHOW tables'), 'read', 'classifySql SHOW should return read');

  // classifySql: DESCRIBE -> read
  assertEqual(iv.classifySql('DESCRIBE t'), 'read', 'classifySql DESCRIBE should return read');

  // classifySql: INSERT -> write
  assertEqual(iv.classifySql('INSERT INTO t VALUES (1)'), 'write', 'classifySql INSERT should return write');

  // classifySql: UPDATE -> write
  assertEqual(iv.classifySql('UPDATE t SET x=1'), 'write', 'classifySql UPDATE should return write');

  // classifySql: DELETE -> write
  assertEqual(iv.classifySql('DELETE FROM t'), 'write', 'classifySql DELETE should return write');

  // classifySql: DROP -> write
  assertEqual(iv.classifySql('DROP TABLE t'), 'write', 'classifySql DROP should return write');

  // classifySql: CREATE -> write
  assertEqual(iv.classifySql('CREATE TABLE t (id INT)'), 'write', 'classifySql CREATE should return write');

  // classifySql: ALTER -> write
  assertEqual(iv.classifySql('ALTER TABLE t ADD col INT'), 'write', 'classifySql ALTER should return write');

  // classifySql: leading whitespace should still classify correctly
  assertEqual(iv.classifySql('  SELECT * FROM t'), 'read', 'classifySql with leading whitespace should return read');

  // isCancel: cancel value -> true
  assertEqual(iv.isCancel('cancel'), true, 'isCancel("cancel") should return true');

  // isCancel: other string -> false
  assertEqual(iv.isCancel('ok'), false, 'isCancel("ok") should return false');

  // isCancel: null -> false
  assertEqual(iv.isCancel(null), false, 'isCancel(null) should return false');

  // isCancel: undefined -> false
  assertEqual(iv.isCancel(undefined), false, 'isCancel(undefined) should return false');

  // isCancel: empty string -> false
  assertEqual(iv.isCancel(''), false, 'isCancel("") should return false');

  console.log('InputValidator example-based tests passed.');

  // --- Property-based tests for QueryBuilder ---
  var qb = loadQueryBuilder();

  /**
   * Feature: sheetdb-rebuild, Property 1: Value quoting round-trip
   * Validates: Requirements 3.3, 3.4, 3.5, 3.6, 3.7, 14.5
   *
   * For any valid JavaScript value (string, number, null, or empty string),
   * calling quoteSqlValue(value) and then parseQuotedValue(result) SHALL
   * produce a value equivalent to the original input.
   *
   * Round-trip semantics:
   * - null → 'NULL' → null
   * - '' (empty string) → 'NULL' → null (empty string collapses to null by design)
   * - numbers → numeric literal → same number
   * - non-empty strings → escaped+quoted → same string
   */
  fc.assert(
    fc.property(
      fc.oneof(
        fc.constant(null),
        fc.constant(''),
        fc.integer(),
        fc.float({ noNaN: true, noDefaultInfinity: true }),
        fc.string()
      ),
      function (value) {
        var quoted = qb.quoteSqlValue(value);
        var parsed = qb.parseQuotedValue(quoted);

        if (value === null || value === '') {
          // null and empty string both collapse to null
          if (parsed !== null) {
            throw new Error(
              'Round-trip failed for ' + JSON.stringify(value) +
              ': expected null but got ' + JSON.stringify(parsed)
            );
          }
        } else if (typeof value === 'number') {
          if (parsed !== value) {
            throw new Error(
              'Round-trip failed for number ' + value +
              ': expected ' + value + ' but got ' + JSON.stringify(parsed)
            );
          }
        } else {
          // non-empty string
          if (parsed !== value) {
            throw new Error(
              'Round-trip failed for string ' + JSON.stringify(value) +
              ': expected ' + JSON.stringify(value) + ' but got ' + JSON.stringify(parsed)
            );
          }
        }
      }
    ),
    { numRuns: 100 }
  );
  console.log('Property 1 (Value quoting round-trip) passed — 100 iterations.');

  /**
   * Feature: sheetdb-rebuild, Property 4: Parameterized query structure
   * Validates: Requirements 3.2, 7.1, 7.3, 8.1, 8.4, 9.1, 9.3
   *
   * For any valid table name, column list, and row of values, the QueryBuilder
   * functions (buildInsertSql, buildUpdateSql, buildDeleteSql) SHALL produce a
   * PreparedQuery where: (a) the SQL string contains ? placeholders and no
   * interpolated user values, and (b) the params array length equals the number
   * of ? placeholders in the SQL string.
   */
  var alphanumCharsArr = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789'.split('');

  var alphanumStringArb = function (minLen, maxLen) {
    return fc.array(fc.constantFrom.apply(fc, alphanumCharsArr), { minLength: minLen, maxLength: maxLen })
      .map(function (arr) { return arr.join(''); });
  };

  var tableNameArb = alphanumStringArb(1, 20);

  var columnNameArb = alphanumStringArb(1, 15);

  var mixedValueArb = fc.oneof(
    // Prefix string values with '@val_' so they are distinguishable from
    // table/column names and SQL keywords in the interpolation check
    fc.string({ minLength: 1, maxLength: 50 }).map(function (s) { return '@val_' + s; }),
    fc.integer({ min: -100000, max: 100000 }),
    fc.constant(null)
  );

  // Generate columns (1-10) and matching values
  var columnsAndValuesArb = fc.integer({ min: 1, max: 10 }).chain(function (colCount) {
    return fc.tuple(
      fc.array(columnNameArb, { minLength: colCount, maxLength: colCount }),
      fc.array(mixedValueArb, { minLength: colCount, maxLength: colCount })
    );
  });

  var countPlaceholders = function (sql) {
    var count = 0;
    for (var i = 0; i < sql.length; i++) {
      if (sql.charAt(i) === '?') count++;
    }
    return count;
  };

  // Check that no user-provided string values appear interpolated in the SQL.
  // String values are generated with a '@val_' prefix so they are clearly
  // distinguishable from table names, column names, and SQL keywords.
  var containsInterpolatedValues = function (sql, values) {
    for (var i = 0; i < values.length; i++) {
      var v = values[i];
      if (typeof v === 'string' && v.length > 0) {
        if (sql.indexOf(v) !== -1) {
          return true;
        }
      }
    }
    return false;
  };

  fc.assert(
    fc.property(
      tableNameArb,
      columnsAndValuesArb,
      function (table, colsAndVals) {
        var columns = colsAndVals[0];
        var values = colsAndVals[1];
        var colCount = columns.length;

        // --- Test buildInsertSql ---
        var insertResult = qb.buildInsertSql(table, columns, values);
        var insertPlaceholderCount = countPlaceholders(insertResult.sql);
        if (insertResult.params.length !== insertPlaceholderCount) {
          throw new Error(
            'buildInsertSql: params.length (' + insertResult.params.length +
            ') !== placeholder count (' + insertPlaceholderCount +
            ') for table=' + table + ', columns=' + JSON.stringify(columns)
          );
        }
        if (containsInterpolatedValues(insertResult.sql, values)) {
          throw new Error(
            'buildInsertSql: SQL contains interpolated user values: ' + insertResult.sql
          );
        }

        // --- Test buildUpdateSql ---
        // Split columns: first half for SET, second half for WHERE (both non-empty)
        if (colCount >= 2) {
          var splitPoint = Math.max(1, Math.floor(colCount / 2));
          var setCols = [];
          var whereColsUpdate = [];
          for (var s = 0; s < splitPoint; s++) { setCols.push(s); }
          for (var w = splitPoint; w < colCount; w++) { whereColsUpdate.push(w); }

          var updateResult = qb.buildUpdateSql(table, columns, values, setCols, whereColsUpdate);
          var updatePlaceholderCount = countPlaceholders(updateResult.sql);
          if (updateResult.params.length !== updatePlaceholderCount) {
            throw new Error(
              'buildUpdateSql: params.length (' + updateResult.params.length +
              ') !== placeholder count (' + updatePlaceholderCount +
              ') for table=' + table + ', columns=' + JSON.stringify(columns)
            );
          }
          if (containsInterpolatedValues(updateResult.sql, values)) {
            throw new Error(
              'buildUpdateSql: SQL contains interpolated user values: ' + updateResult.sql
            );
          }
        }

        // --- Test buildDeleteSql ---
        var whereColsDelete = [];
        for (var d = 0; d < colCount; d++) { whereColsDelete.push(d); }

        var deleteResult = qb.buildDeleteSql(table, columns, values, whereColsDelete);
        var deletePlaceholderCount = countPlaceholders(deleteResult.sql);
        if (deleteResult.params.length !== deletePlaceholderCount) {
          throw new Error(
            'buildDeleteSql: params.length (' + deleteResult.params.length +
            ') !== placeholder count (' + deletePlaceholderCount +
            ') for table=' + table + ', columns=' + JSON.stringify(columns)
          );
        }
        if (containsInterpolatedValues(deleteResult.sql, values)) {
          throw new Error(
            'buildDeleteSql: SQL contains interpolated user values: ' + deleteResult.sql
          );
        }
      }
    ),
    { numRuns: 100 }
  );
  console.log('Property 4 (Parameterized query structure) passed — 100 iterations.');

  // --- Example-based unit tests for QueryBuilder ---

  // quoteSqlValue: null → 'NULL'
  assertEqual(qb.quoteSqlValue(null), 'NULL', 'quoteSqlValue(null) should return NULL');

  // quoteSqlValue: empty string → 'NULL'
  assertEqual(qb.quoteSqlValue(''), 'NULL', 'quoteSqlValue("") should return NULL');

  // quoteSqlValue: string with single quotes → escaped
  assertEqual(qb.quoteSqlValue("O'Reilly"), "'O''Reilly'", "quoteSqlValue should escape single quotes");

  // quoteSqlValue: zero → '0'
  assertEqual(qb.quoteSqlValue(0), '0', 'quoteSqlValue(0) should return 0');

  // quoteSqlValue: negative number → '-42'
  assertEqual(qb.quoteSqlValue(-42), '-42', 'quoteSqlValue(-42) should return -42');

  // quoteSqlValue: float → '3.14'
  assertEqual(qb.quoteSqlValue(3.14), '3.14', 'quoteSqlValue(3.14) should return 3.14');

  // quoteSqlValue: regular string → quoted
  assertEqual(qb.quoteSqlValue('hello'), "'hello'", "quoteSqlValue('hello') should return quoted string");

  // quoteSqlValue: Date with custom formatter → quoted formatted string
  assertEqual(
    qb.quoteSqlValue(new Date(2026, 0, 2, 3, 4, 5), function (d) { return '2026-01-02 03:04:05'; }),
    "'2026-01-02 03:04:05'",
    'quoteSqlValue(Date) should use formatDateFn and return quoted result'
  );

  // buildInsertSql: single column
  var insertSingle = qb.buildInsertSql('users', ['name'], ['Alice']);
  assertEqual(insertSingle.sql, 'INSERT INTO users (name) VALUES (?)', 'buildInsertSql single column SQL');
  assertEqual(JSON.stringify(insertSingle.params), JSON.stringify(['Alice']), 'buildInsertSql single column params');

  // buildInsertSql: multiple columns
  var insertMulti = qb.buildInsertSql('users', ['name', 'age'], ['Alice', 30]);
  assertEqual(insertMulti.sql, 'INSERT INTO users (name,age) VALUES (?,?)', 'buildInsertSql multiple columns SQL');
  assertEqual(JSON.stringify(insertMulti.params), JSON.stringify(['Alice', 30]), 'buildInsertSql multiple columns params');

  // buildUpdateSql: SET name,age WHERE id
  var updateResult = qb.buildUpdateSql('users', ['name', 'age', 'id'], ['Alice', 30, 1], [0, 1], [2]);
  assertEqual(updateResult.sql, 'UPDATE users SET name=?,age=? WHERE id=?', 'buildUpdateSql SQL');
  assertEqual(JSON.stringify(updateResult.params), JSON.stringify(['Alice', 30, 1]), 'buildUpdateSql params');

  // buildDeleteSql: single WHERE column
  var deleteSingle = qb.buildDeleteSql('users', ['id'], [1], [0]);
  assertEqual(deleteSingle.sql, 'DELETE FROM users WHERE id=?', 'buildDeleteSql single WHERE SQL');
  assertEqual(JSON.stringify(deleteSingle.params), JSON.stringify([1]), 'buildDeleteSql single WHERE params');

  // buildDeleteSql: multiple WHERE columns
  var deleteMulti = qb.buildDeleteSql('users', ['name', 'age'], ['Alice', 30], [0, 1]);
  assertEqual(deleteMulti.sql, 'DELETE FROM users WHERE name=? AND age=?', 'buildDeleteSql multiple WHERE SQL');
  assertEqual(JSON.stringify(deleteMulti.params), JSON.stringify(['Alice', 30]), 'buildDeleteSql multiple WHERE params');

  console.log('QueryBuilder example-based tests passed.');

  // --- Property-based tests for ConfigManager ---
  var cm = loadConfigManager();

  /**
   * Feature: sheetdb-rebuild, Property 5: Config validation
   * Validates: Requirements 2.3, 2.5
   *
   * For any config object with url, user, and password fields,
   * hasValidConfig(config) SHALL return true if and only if all three fields
   * are non-empty strings. If any field is empty, null, undefined, or not a
   * string, it SHALL return false.
   */
  var fieldValueArb = fc.oneof(
    fc.string({ minLength: 1, maxLength: 50 }),  // non-empty string
    fc.constant(''),                              // empty string
    fc.constant(null),                            // null
    fc.constant(undefined),                       // undefined
    fc.integer({ min: -1000, max: 1000 }),        // number
    fc.boolean()                                  // boolean
  );

  fc.assert(
    fc.property(
      fieldValueArb,
      fieldValueArb,
      fieldValueArb,
      function (url, user, password) {
        var config = { url: url, user: user, password: password };
        var result = cm.hasValidConfig(config);

        var allNonEmptyStrings =
          typeof url === 'string' && url !== '' &&
          typeof user === 'string' && user !== '' &&
          typeof password === 'string' && password !== '';

        if (result !== allNonEmptyStrings) {
          throw new Error(
            'hasValidConfig(' + JSON.stringify(config) + ') returned ' + result +
            ' but expected ' + allNonEmptyStrings
          );
        }
      }
    ),
    { numRuns: 100 }
  );
  console.log('Property 5 (Config validation) passed — 100 iterations.');

  // --- Example-based unit tests for hasValidConfig ---

  // All valid fields → true
  assertEqual(
    cm.hasValidConfig({ url: 'jdbc:mysql://host', user: 'root', password: 'pass' }),
    true,
    'hasValidConfig: all valid fields should return true'
  );

  // Empty url → false
  assertEqual(
    cm.hasValidConfig({ url: '', user: 'root', password: 'pass' }),
    false,
    'hasValidConfig: empty url should return false'
  );

  // Empty user → false
  assertEqual(
    cm.hasValidConfig({ url: 'jdbc:mysql://host', user: '', password: 'pass' }),
    false,
    'hasValidConfig: empty user should return false'
  );

  // Empty password → false
  assertEqual(
    cm.hasValidConfig({ url: 'jdbc:mysql://host', user: 'root', password: '' }),
    false,
    'hasValidConfig: empty password should return false'
  );

  // Null url → false
  assertEqual(
    cm.hasValidConfig({ url: null, user: 'root', password: 'pass' }),
    false,
    'hasValidConfig: null url should return false'
  );

  // Undefined user → false
  assertEqual(
    cm.hasValidConfig({ url: 'jdbc:mysql://host', user: undefined, password: 'pass' }),
    false,
    'hasValidConfig: undefined user should return false'
  );

  // Number password → false
  assertEqual(
    cm.hasValidConfig({ url: 'jdbc:mysql://host', user: 'root', password: 123 }),
    false,
    'hasValidConfig: number password should return false'
  );

  // Boolean url → false
  assertEqual(
    cm.hasValidConfig({ url: true, user: 'root', password: 'pass' }),
    false,
    'hasValidConfig: boolean url should return false'
  );

  console.log('ConfigManager (hasValidConfig) example-based tests passed.');

  // --- Property-based tests for BulkReport (TableOps) ---
  var to = loadTableOps();

  /**
   * Feature: sheetdb-rebuild, Property 6: Bulk report accuracy
   * Validates: Requirements 11.3
   *
   * For any sequence of bulk operation outcomes (each either success or failure),
   * the resulting BulkReport SHALL have total === succeeded + failed,
   * succeeded equal to the count of successes, and failed equal to the count of failures.
   */
  fc.assert(
    fc.property(
      fc.array(fc.boolean(), { minLength: 0, maxLength: 100 }),
      function (outcomes) {
        var report = to.createBulkReport(outcomes.length);

        for (var i = 0; i < outcomes.length; i++) {
          to.recordBulkResult(report, outcomes[i], 'error msg');
        }

        var expectedSucceeded = outcomes.filter(function (b) { return b === true; }).length;
        var expectedFailed = outcomes.filter(function (b) { return b === false; }).length;

        if (report.total !== outcomes.length) {
          throw new Error(
            'report.total (' + report.total + ') !== outcomes.length (' + outcomes.length + ')'
          );
        }
        if (report.succeeded !== expectedSucceeded) {
          throw new Error(
            'report.succeeded (' + report.succeeded + ') !== expected (' + expectedSucceeded + ')'
          );
        }
        if (report.failed !== expectedFailed) {
          throw new Error(
            'report.failed (' + report.failed + ') !== expected (' + expectedFailed + ')'
          );
        }
        if (report.total !== report.succeeded + report.failed) {
          throw new Error(
            'report.total (' + report.total + ') !== succeeded + failed (' +
            (report.succeeded + report.failed) + ')'
          );
        }
      }
    ),
    { numRuns: 100 }
  );
  console.log('Property 6 (Bulk report accuracy) passed — 100 iterations.');

  // --- Example-based unit tests for BulkReport logic ---

  // All successes (3 rows)
  var reportAllSuccess = to.createBulkReport(3);
  to.recordBulkResult(reportAllSuccess, true);
  to.recordBulkResult(reportAllSuccess, true);
  to.recordBulkResult(reportAllSuccess, true);
  assertEqual(reportAllSuccess.total, 3, 'BulkReport all successes: total should be 3');
  assertEqual(reportAllSuccess.succeeded, 3, 'BulkReport all successes: succeeded should be 3');
  assertEqual(reportAllSuccess.failed, 0, 'BulkReport all successes: failed should be 0');
  assertEqual(JSON.stringify(reportAllSuccess.errors), JSON.stringify([]), 'BulkReport all successes: errors should be empty');

  // All failures (3 rows)
  var reportAllFail = to.createBulkReport(3);
  to.recordBulkResult(reportAllFail, false, 'err1');
  to.recordBulkResult(reportAllFail, false, 'err2');
  to.recordBulkResult(reportAllFail, false, 'err3');
  assertEqual(reportAllFail.total, 3, 'BulkReport all failures: total should be 3');
  assertEqual(reportAllFail.succeeded, 0, 'BulkReport all failures: succeeded should be 0');
  assertEqual(reportAllFail.failed, 3, 'BulkReport all failures: failed should be 3');
  assertEqual(reportAllFail.errors.length, 3, 'BulkReport all failures: errors.length should be 3');

  // Mixed results (4 rows): success, failure, success, failure
  var reportMixed = to.createBulkReport(4);
  to.recordBulkResult(reportMixed, true);
  to.recordBulkResult(reportMixed, false, 'fail1');
  to.recordBulkResult(reportMixed, true);
  to.recordBulkResult(reportMixed, false, 'fail2');
  assertEqual(reportMixed.total, 4, 'BulkReport mixed: total should be 4');
  assertEqual(reportMixed.succeeded, 2, 'BulkReport mixed: succeeded should be 2');
  assertEqual(reportMixed.failed, 2, 'BulkReport mixed: failed should be 2');
  assertEqual(reportMixed.errors.length, 2, 'BulkReport mixed: errors.length should be 2');

  // Empty array (0 rows)
  var reportEmpty = to.createBulkReport(0);
  assertEqual(reportEmpty.total, 0, 'BulkReport empty: total should be 0');
  assertEqual(reportEmpty.succeeded, 0, 'BulkReport empty: succeeded should be 0');
  assertEqual(reportEmpty.failed, 0, 'BulkReport empty: failed should be 0');
  assertEqual(reportEmpty.errors.length, 0, 'BulkReport empty: errors.length should be 0');

  console.log('BulkReport example-based tests passed.');

  console.log('All tests passed.');
})();
