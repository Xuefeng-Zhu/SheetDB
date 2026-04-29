# Implementation Plan: SheetDB Rebuild

## Overview

Rebuild the SheetDB Google Sheets Add-on from a monolithic codebase into a modular architecture with 11 focused modules. Implementation proceeds bottom-up: foundation modules with no dependencies first, then core services, then high-level orchestration, and finally integration wiring. All pure-logic modules are tested with fast-check property-based tests and example-based unit tests under Node.js.

## Tasks

- [x] 1. Set up project foundation and Constants module
  - [x] 1.1 Install fast-check as a dev dependency and update package.json test script
    - Run `npm install --save-dev fast-check`
    - Ensure `npm test` runs `node tests/run-tests.js`
    - _Requirements: 14.1_

  - [x] 1.2 Create `src/Constants.gs` with all shared constants
    - Define sheet name constants: `SHEET_NAME_CONFIG = 'Configue'`, `SHEET_NAME_SQL = 'SQL'`
    - Define all menu label constants: `MENU_SQL`, `MENU_TABLE`, `MENU_SHOW_PROMPT`, `MENU_REFRESH`, `MENU_CLEAR_HISTORY`, `MENU_CONFIGURE`, `MENU_LOAD_TABLES`, `MENU_REFRESH_TABLE`, `MENU_DROP_TABLE`, `MENU_FAST_INSERT`, `MENU_FAST_UPDATE`, `MENU_FAST_DELETE`
    - Define JDBC constant: `JDBC_TIMEOUT_SECONDS = 30`
    - Define UI constants: `CANCEL_BUTTON_VALUE = 'cancel'`, `MSG_CONFIGURE_FIRST = 'Please configure the system first!'`
    - _Requirements: 10.3, 13.2_

- [x] 2. Implement InputValidator module
  - [x] 2.1 Create `src/InputValidator.gs` with pure validation functions
    - Implement `validateSql(input)`: trims input, throws descriptive error if empty or whitespace-only, returns trimmed SQL
    - Implement `classifySql(sql)`: returns `'read'` for SELECT/SHOW/DESCRIBE (case-insensitive first keyword), `'write'` for everything else
    - Implement `isCancel(value)`: returns `true` if value equals `CANCEL_BUTTON_VALUE`
    - No Apps Script dependencies — pure functions only
    - _Requirements: 3.1, 13.3, 13.4_

  - [x] 2.2 Write property test: Whitespace SQL rejection (Property 2)
    - **Property 2: Whitespace SQL rejection**
    - Generate strings composed entirely of whitespace characters (space, tab, \n, \r) of varying lengths
    - Assert `validateSql(input)` always throws a descriptive error and never returns a value
    - **Validates: Requirements 3.1**

  - [x] 2.3 Write property test: SQL classification correctness (Property 3)
    - **Property 3: SQL classification correctness**
    - Generate SQL strings with random first keywords from read set (SELECT, SHOW, DESCRIBE) and write set (INSERT, UPDATE, DELETE, DROP, CREATE, ALTER), with random case and leading whitespace
    - Assert `classifySql` returns `'read'` for read keywords and `'write'` for write keywords
    - **Validates: Requirements 14.4**

  - [x] 2.4 Write example-based unit tests for InputValidator
    - Test `validateSql` with empty string, whitespace-only, valid SQL
    - Test `classifySql` with each keyword variant, case sensitivity, leading whitespace
    - Test `isCancel` with cancel value, other values, null, undefined
    - _Requirements: 14.3, 14.4_

- [x] 3. Implement QueryBuilder module
  - [x] 3.1 Create `src/QueryBuilder.gs` with SQL construction functions
    - Implement `quoteSqlValue(value, formatDateFn?)`: NULL for null/empty, numeric literal for numbers, escaped+quoted for strings, formatted+quoted for Dates. Accept optional `formatDateFn` parameter for testability
    - Implement `parseQuotedValue(quoted)`: inverse of `quoteSqlValue` — parses SQL literal back to JS value for round-trip testing
    - Implement `buildInsertSql(table, columns, values)`: returns `{sql, params}` with `?` placeholders and param array
    - Implement `buildUpdateSql(table, columns, values, setCols, whereCols)`: returns `{sql, params}` with SET clause from `setCols` indices and WHERE clause from `whereCols` indices
    - Implement `buildDeleteSql(table, columns, values, whereCols)`: returns `{sql, params}` with WHERE clause from specified columns
    - No Apps Script dependencies — pure functions only (Date formatting injected via parameter)
    - _Requirements: 3.2, 3.3, 3.4, 3.5, 3.6, 3.7, 13.3, 13.4_

  - [x] 3.2 Write property test: Value quoting round-trip (Property 1)
    - **Property 1: Value quoting round-trip**
    - Generate strings (with special chars, quotes, unicode), numbers (int, float, negative, zero), null, empty string, Date objects
    - Assert `parseQuotedValue(quoteSqlValue(value))` produces an equivalent value to the original input
    - **Validates: Requirements 3.3, 3.4, 3.5, 3.6, 3.7, 14.5**

  - [x] 3.3 Write property test: Parameterized query structure (Property 4)
    - **Property 4: Parameterized query structure**
    - Generate random table names (alphanumeric), column name arrays, value arrays of mixed types
    - Assert for `buildInsertSql`, `buildUpdateSql`, `buildDeleteSql`: (a) SQL contains `?` placeholders and no interpolated user values, (b) `params.length` equals the count of `?` in the SQL string
    - **Validates: Requirements 3.2, 7.1, 7.3, 8.1, 8.4, 9.1, 9.3**

  - [x] 3.4 Write example-based unit tests for QueryBuilder
    - Test `quoteSqlValue` with null, empty string, string with single quotes, Date, zero, negative numbers, regular strings
    - Test `buildInsertSql` with single column, multiple columns, mixed value types
    - Test `buildUpdateSql` with various SET/WHERE column splits
    - Test `buildDeleteSql` with single and multiple WHERE columns
    - _Requirements: 14.2, 14.5_

- [x] 4. Checkpoint — Verify foundation modules
  - Ensure all tests pass, ask the user if questions arise.

- [x] 5. Implement ConfigManager module
  - [x] 5.1 Create `src/ConfigManager.gs` with configuration management functions
    - Implement `getConfigSheet()`: returns Config_Sheet, creating it with labeled rows (URL, Admin, Password) if absent. Use `SHEET_NAME_CONFIG` constant
    - Implement `getConfig()`: reads credentials from Config_Sheet, returns `{url, user, password}`
    - Implement `hasValidConfig(config)`: returns `true` if all three fields are non-empty strings
    - Implement `saveConfig(url, user, password)`: writes credentials to Config_Sheet
    - Implement `runConfigDialog()`: prompts user for URL, username, password in sequence via `Browser.inputBox`. Returns `false` if cancelled at any step, displays cancellation message. On success, calls `saveConfig`
    - _Requirements: 2.1, 2.2, 2.3, 2.4, 2.5, 2.6_

  - [x] 5.2 Write property test: Config validation (Property 5)
    - **Property 5: Config validation**
    - Generate config objects with random combinations of empty/non-empty/null/undefined string fields for url, user, password
    - Assert `hasValidConfig(config)` returns `true` if and only if all three fields are non-empty strings
    - **Validates: Requirements 2.3, 2.5**

  - [x] 5.3 Write example-based unit tests for hasValidConfig
    - Test with all valid fields, each field empty, null, undefined, non-string types
    - _Requirements: 2.3, 2.5_

- [x] 6. Implement JdbcService module
  - [x] 6.1 Create `src/JdbcService.gs` with unified JDBC execution functions
    - Implement `getConnection_()`: internal function that gets config via `getConfig()`, opens JDBC connection, sets statement timeout to `JDBC_TIMEOUT_SECONDS`. Throws with context on failure
    - Implement `executeRead(sql)`: opens connection, creates Statement, executes query, returns array of row arrays. Closes ResultSet, Statement, Connection in finally block. Throws error with original message + failing SQL
    - Implement `executeWrite(sql)`: opens connection, creates Statement, executes update, returns affected row count. Closes Statement, Connection in finally block. Throws error with original message + failing SQL
    - Implement `executePrepared(sql, params)`: opens connection, creates PreparedStatement, binds params by type, executes update, returns affected row count. Closes all resources in finally block
    - All close errors in finally blocks are suppressed (logged to `console.error`) to avoid masking original errors
    - _Requirements: 1.1, 1.2, 1.3, 1.4, 1.5, 1.6, 11.4_

- [x] 7. Implement SchemaCache module
  - [x] 7.1 Create `src/SchemaCache.gs` with in-memory schema caching
    - Define module-level `schemaCache_` variable (reset each script execution)
    - Implement `getColumns(tableName)`: returns cached column names, or fetches via `JdbcService.executeRead('DESCRIBE ' + tableName)` and caches
    - Implement `invalidate(tableName?)`: clears cache for one table or all tables
    - _Requirements: 5.4_

- [x] 8. Implement Helpers module
  - [x] 8.1 Create `src/Helpers.gs` with shared utility functions
    - Implement `ensureSheet(name)`: returns existing sheet by name or creates a new one
    - Implement `ensureSqlSheet()`: shorthand for `ensureSheet(SHEET_NAME_SQL)`
    - Implement `appendSqlHistory(rows)`: appends rows to the SQL history sheet
    - Implement `requireConfig()`: gets config via `getConfig()`, checks `hasValidConfig()`, shows `MSG_CONFIGURE_FIRST` error dialog and throws if not configured, otherwise returns config
    - _Requirements: 11.2, 13.1_

- [x] 9. Implement ResultRenderer module
  - [x] 9.1 Create `src/ResultRenderer.gs` with spreadsheet result rendering functions
    - Implement `renderReadResults(sheet, sql, rows)`: writes status row (sql + " Success"), then data rows with one row per result row. Appends blank separator row
    - Implement `renderWriteResult(sheet, sql)`: writes status row (sql + " Success") and blank separator
    - Implement `renderNoResults(sheet, sql)`: writes status row indicating zero results returned
    - Implement `replaceResultBlock(sheet, startRow, newRows)`: for Refresh — finds old result block boundaries (delimited by blank row or end of data), deletes old rows, inserts new rows in place
    - _Requirements: 4.3, 12.1, 12.2, 12.3, 12.4_

- [x] 10. Checkpoint — Verify core services
  - Ensure all tests pass, ask the user if questions arise.

- [x] 11. Implement SqlPrompt module
  - [x] 11.1 Create `src/SqlPrompt.gs` with SQL prompt, refresh, and history functions
    - Implement `showPrompt()`: displays input dialog via `Browser.inputBox`, validates input with `validateSql`, classifies with `classifySql`, executes via `executeRead` or `executeWrite`, renders results via `ResultRenderer`, records history via `appendSqlHistory`. Handles cancel with no action
    - Implement `refresh()`: reads SQL from selected cell, strips " Success"/" Failed" suffix, re-executes, replaces result block in place via `replaceResultBlock`
    - Implement `clearHistory()`: confirms with user via `Browser.msgBox`, clears active sheet on confirmation
    - _Requirements: 4.1, 4.2, 4.3, 4.4, 4.5, 4.6, 4.7_

- [x] 12. Implement TableOps module
  - [x] 12.1 Create `src/TableOps.gs` with table loading, refresh, and drop functions
    - Implement `loadTables()`: fetches table list via `executeRead('SHOW tables')`, creates or refreshes a Table_Sheet for each table using `SchemaCache.getColumns` for headers and `executeRead('SELECT * FROM ...')` for data. Bold headers in row 1
    - Implement `refreshTable()`: if on SQL_Sheet, prompts for table name; otherwise uses current sheet name. Clears and reloads schema + data. Shows error if table sheet not found
    - Implement `dropTable()`: if on SQL_Sheet, prompts for table name. Executes `DROP TABLE` via `executeWrite`, records in SQL history, removes sheet. Handles cancel and errors
    - _Requirements: 5.1, 5.2, 5.3, 5.4, 5.5, 5.6, 5.7, 6.1, 6.2, 6.3, 6.4, 6.5_

  - [x] 12.2 Create bulk operation functions in `src/TableOps.gs`
    - Implement `fastInsert()`: if on SQL_Sheet, prompts for table name and optional column list; otherwise uses Table_Sheet headers. Reads selected range, builds parameterized INSERT per row via `buildInsertSql`, executes via `executePrepared`, records history. Continue-on-error with BulkReport
    - Implement `fastUpdate()`: if on SQL_Sheet, prompts for table name and columns; otherwise uses Table_Sheet headers. Shows checkbox dialog for column selection (SET vs WHERE). Builds parameterized UPDATE per row via `buildUpdateSql`, executes via `executePrepared`. Continue-on-error with BulkReport
    - Implement `fastDelete()`: if on SQL_Sheet, prompts for table name and columns; otherwise uses Table_Sheet headers. Builds parameterized DELETE per row via `buildDeleteSql`, executes via `executePrepared`, removes sheet rows on success. Continue-on-error with BulkReport
    - All bulk operations report success/failure counts via `Browser.msgBox` on completion
    - _Requirements: 7.1, 7.2, 7.3, 7.4, 7.5, 7.6, 8.1, 8.2, 8.3, 8.4, 8.5, 8.6, 8.7, 9.1, 9.2, 9.3, 9.4, 9.5, 9.6, 11.3_

  - [x] 12.3 Write property test: Bulk report accuracy (Property 6)
    - **Property 6: Bulk report accuracy**
    - Generate random boolean arrays representing success/failure sequences
    - Assert the resulting BulkReport has `total === succeeded + failed`, `succeeded` equals count of true values, `failed` equals count of false values
    - **Validates: Requirements 11.3**

  - [x] 12.4 Write example-based unit tests for BulkReport logic
    - Test with all successes, all failures, mixed results, empty array
    - _Requirements: 11.3_

- [x] 13. Implement MenuController and wire entry points
  - [x] 13.1 Create `src/MenuController.gs` with menu registration
    - Implement `onOpen()`: registers "SQL" menu with items (Show Prompt, Refresh, Clear History, Configure) and "Table" menu with items (Load Tables, Refresh Table, Drop Table, Fast Insert, Fast Update, Fast Delete)
    - Use constant values from `Constants.gs` for all menu labels
    - Route each menu item to the appropriate handler function in SqlPrompt, TableOps, or ConfigManager
    - _Requirements: 10.1, 10.2, 10.3_

  - [x] 13.2 Create `src/Trigger.gs` with the onEdit trigger
    - Preserve the existing `onEdit` trigger behavior or update as needed for the new architecture
    - _Requirements: 8.2_

- [x] 14. Checkpoint — Verify full integration
  - Ensure all tests pass, ask the user if questions arise.

- [x] 15. Remove legacy source files and finalize
  - [x] 15.1 Remove old monolithic source files
    - Delete `src/Code.gs` (replaced by MenuController, SqlPrompt, ResultRenderer)
    - Delete `src/Config.gs` (replaced by ConfigManager)
    - Delete `src/Table.gs` (replaced by TableOps, SchemaCache)
    - Delete `src/Utils.gs` (replaced by Constants, InputValidator, QueryBuilder, Helpers)
    - Verify no references to deleted files remain
    - _Requirements: 13.1_

  - [x] 15.2 Update test runner to cover all new modules
    - Update `tests/run-tests.js` to load and test all pure modules: InputValidator, QueryBuilder, ConfigManager (hasValidConfig), BulkReport logic
    - Include both property-based tests (fast-check) and example-based unit tests
    - Ensure `npm test` passes cleanly
    - _Requirements: 14.1, 14.2, 14.3, 14.4, 14.5_

- [x] 16. Final checkpoint — Ensure all tests pass
  - Ensure all tests pass, ask the user if questions arise.

## Notes

- Tasks marked with `*` are optional and can be skipped for faster MVP
- Each task references specific requirements for traceability
- Checkpoints ensure incremental validation after each major phase
- Property tests validate the 6 universal correctness properties defined in the design
- Unit tests validate specific examples and edge cases
- All pure modules (InputValidator, QueryBuilder, hasValidConfig, BulkReport) are testable under Node.js without Apps Script stubs
- The `quoteSqlValue` function accepts an optional `formatDateFn` parameter so tests can inject a deterministic Date formatter
