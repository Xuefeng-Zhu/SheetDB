# Requirements Document

## Introduction

SheetDB is a Google Sheets Add-on that connects spreadsheets to remote SQL databases via JDBC. This rebuild addresses architectural debt, code duplication, fragile conventions, missing error handling, and limited testability in the current implementation. The goal is to produce a cleaner, more robust, and more maintainable codebase while preserving the core value proposition: letting users query and manipulate SQL databases directly from Google Sheets.

## Glossary

- **Add_On**: The SheetDB Google Sheets Add-on running in the Apps Script V8 runtime
- **JDBC_Service**: The unified module responsible for all JDBC connection management and SQL execution
- **Config_Manager**: The module responsible for reading, writing, and validating database credentials
- **Config_Sheet**: The hidden spreadsheet sheet storing database connection credentials (named "Configue" for backward compatibility)
- **SQL_Sheet**: The dedicated spreadsheet sheet that records SQL execution history
- **Table_Sheet**: A spreadsheet sheet representing a single database table, with column headers in row 1 and data rows below
- **Query_Builder**: The module responsible for constructing parameterized SQL statements from user input and cell data
- **Menu_Controller**: The module that registers custom menus and routes menu actions to the appropriate handlers
- **Schema_Cache**: An in-memory cache of table column metadata retrieved from the database, valid for the duration of a script execution
- **Bulk_Operation**: An insert, update, or delete operation applied to multiple rows selected in a spreadsheet range
- **User**: A person interacting with the Add_On through Google Sheets menus and dialogs
- **Result_Renderer**: The module responsible for writing SQL query results into spreadsheet cells
- **Input_Validator**: The module responsible for sanitizing and validating user-provided SQL and configuration values

## Requirements

### Requirement 1: Unified JDBC Execution

**User Story:** As a developer, I want a single JDBC execution module, so that connection management and query execution logic is not duplicated across files.

#### Acceptance Criteria

1. THE JDBC_Service SHALL provide a single function for executing read queries that returns an array of row arrays
2. THE JDBC_Service SHALL provide a single function for executing write statements that returns the number of affected rows
3. THE JDBC_Service SHALL close all JDBC resources (Connection, Statement, ResultSet) in a finally block after every execution
4. WHEN the JDBC_Service opens a connection, THE JDBC_Service SHALL set a statement timeout of 30 seconds
5. IF a JDBC connection fails, THEN THE JDBC_Service SHALL throw an error containing the original error message and the failing SQL statement
6. IF a JDBC statement execution fails, THEN THE JDBC_Service SHALL throw an error containing the original error message and the failing SQL statement

### Requirement 2: Configuration Management

**User Story:** As a user, I want to configure my database credentials once and have them stored securely, so that I can connect to my database without re-entering credentials each time.

#### Acceptance Criteria

1. THE Config_Manager SHALL store credentials (URL, username, password) in the Config_Sheet
2. THE Config_Manager SHALL read the Config_Sheet named "Configue" for backward compatibility with existing deployments
3. WHEN a User provides new credentials via the configure dialog, THE Config_Manager SHALL validate that the URL, username, and password fields are non-empty before saving
4. IF the Config_Sheet does not exist, THEN THE Config_Manager SHALL create the Config_Sheet with labeled rows for URL, Admin, and Password
5. THE Config_Manager SHALL provide a function that returns a boolean indicating whether valid credentials are currently stored
6. IF the User cancels any step of the configuration dialog, THEN THE Config_Manager SHALL abort the configuration process and display a cancellation message

### Requirement 3: Input Validation and SQL Safety

**User Story:** As a developer, I want all user-provided SQL and values to be validated and safely quoted, so that malformed input does not cause unexpected database operations.

#### Acceptance Criteria

1. THE Input_Validator SHALL reject SQL statements that are empty or contain only whitespace by returning a descriptive error
2. THE Query_Builder SHALL use parameterized queries (PreparedStatement) for all bulk operations (insert, update, delete) instead of string concatenation
3. THE Query_Builder SHALL quote string values by escaping single quotes (replacing `'` with `''`) and wrapping in single quotes
4. THE Query_Builder SHALL represent null values and empty strings as the SQL keyword NULL
5. THE Query_Builder SHALL represent numeric values as unquoted numeric literals
6. THE Query_Builder SHALL format Date values as quoted strings in "yyyy-MM-dd HH:mm:ss" format using the script timezone
7. FOR ALL valid values, quoting a value and then parsing the quoted representation SHALL produce an equivalent value (round-trip property)

### Requirement 4: SQL Prompt and History

**User Story:** As a user, I want to execute arbitrary SQL statements from a prompt and have them recorded in a history sheet, so that I can review and re-run past queries.

#### Acceptance Criteria

1. WHEN the User selects "Show Prompt" from the SQL menu, THE Add_On SHALL display an input dialog for entering a SQL statement
2. WHEN the User submits a valid SQL statement, THE Add_On SHALL execute the statement via the JDBC_Service and display results on the active sheet
3. WHEN a SELECT, SHOW, or DESCRIBE statement returns results, THE Result_Renderer SHALL write column headers in the first result row and data rows below
4. WHEN a write statement (INSERT, UPDATE, DELETE, DROP, CREATE, ALTER) executes successfully, THE Add_On SHALL append a success entry to the SQL_Sheet
5. IF the User cancels the SQL prompt dialog, THEN THE Add_On SHALL take no further action
6. WHEN the User selects "Refresh" from the SQL menu, THE Add_On SHALL re-execute the SQL statement found in the currently selected cell and replace the previous result block in place
7. WHEN the User selects "Clear History" from the SQL menu and confirms, THE Add_On SHALL clear all content from the active sheet

### Requirement 5: Table Loading and Schema Management

**User Story:** As a user, I want to load all database tables into individual sheets with their schema and data, so that I can browse and work with table contents in the spreadsheet.

#### Acceptance Criteria

1. WHEN the User selects "Load Tables", THE Add_On SHALL retrieve the list of tables from the database and create or refresh a Table_Sheet for each table
2. WHEN creating a new Table_Sheet, THE Add_On SHALL write column names as bold headers in row 1 and data rows starting from row 2
3. WHEN a Table_Sheet already exists for a table, THE Add_On SHALL clear and reload the schema and data for that sheet
4. THE Schema_Cache SHALL cache column metadata per table for the duration of a single script execution to avoid redundant DESCRIBE queries
5. WHEN the User selects "Refresh Table" while on a Table_Sheet, THE Add_On SHALL reload the schema and data for that table
6. WHEN the User selects "Refresh Table" while on the SQL_Sheet, THE Add_On SHALL prompt for a table name and refresh the corresponding Table_Sheet
7. IF the specified table name does not correspond to an existing Table_Sheet, THEN THE Add_On SHALL display an error message

### Requirement 6: Table Drop

**User Story:** As a user, I want to drop a database table and remove its sheet, so that I can manage my database schema from the spreadsheet.

#### Acceptance Criteria

1. WHEN the User selects "Drop Table" while on a Table_Sheet, THE Add_On SHALL execute a DROP TABLE statement for that table and remove the Table_Sheet
2. WHEN the User selects "Drop Table" while on the SQL_Sheet, THE Add_On SHALL prompt for a table name before dropping
3. IF the User cancels the drop table dialog, THEN THE Add_On SHALL take no further action
4. WHEN a table is dropped successfully, THE Add_On SHALL record the DROP statement in the SQL_Sheet history
5. IF the DROP TABLE statement fails, THEN THE Add_On SHALL display the error message to the User and retain the Table_Sheet

### Requirement 7: Bulk Insert

**User Story:** As a user, I want to insert multiple rows into a database table by selecting cell ranges, so that I can quickly add data without writing SQL manually.

#### Acceptance Criteria

1. WHEN the User selects "Fast Insert" while on a Table_Sheet, THE Query_Builder SHALL construct INSERT statements using the Table_Sheet column headers and the selected cell range values
2. WHEN the User selects "Fast Insert" while on the SQL_Sheet, THE Add_On SHALL prompt for a table name and optional column list
3. THE Query_Builder SHALL generate one INSERT statement per selected row using parameterized queries
4. WHEN all rows are inserted successfully, THE Add_On SHALL record each INSERT statement in the SQL_Sheet history
5. IF any INSERT statement fails, THEN THE Add_On SHALL display the error message and continue processing remaining rows
6. IF the User cancels the fast insert dialog, THEN THE Add_On SHALL take no further action

### Requirement 8: Bulk Update

**User Story:** As a user, I want to update multiple rows in a database table by selecting modified cells, so that I can push spreadsheet edits back to the database.

#### Acceptance Criteria

1. WHEN the User selects "Fast Update" while on a Table_Sheet, THE Query_Builder SHALL construct UPDATE statements using the Table_Sheet column headers, the selected cell range values, and a column-selection mechanism to identify which columns to update
2. THE Add_On SHALL use a checkbox-based or dialog-based column selection mechanism instead of the bold-cell formatting convention
3. WHEN the User selects "Fast Update" while on the SQL_Sheet, THE Add_On SHALL prompt for a table name and column list
4. THE Query_Builder SHALL use non-selected columns as WHERE clause conditions and selected columns as SET clause values
5. WHEN all rows are updated successfully, THE Add_On SHALL record each UPDATE statement in the SQL_Sheet history
6. IF any UPDATE statement fails, THEN THE Add_On SHALL display the error message and continue processing remaining rows
7. IF the User cancels the fast update dialog, THEN THE Add_On SHALL take no further action

### Requirement 9: Bulk Delete

**User Story:** As a user, I want to delete multiple rows from a database table by selecting cell ranges, so that I can remove data without writing SQL manually.

#### Acceptance Criteria

1. WHEN the User selects "Fast Delete" while on a Table_Sheet, THE Query_Builder SHALL construct DELETE statements using the Table_Sheet column headers and the selected cell range values as WHERE clause conditions
2. WHEN the User selects "Fast Delete" while on the SQL_Sheet, THE Add_On SHALL prompt for a table name and column list
3. THE Query_Builder SHALL generate one DELETE statement per selected row using parameterized queries
4. WHEN all rows are deleted successfully, THE Add_On SHALL record each DELETE statement in the SQL_Sheet history and remove the corresponding rows from the Table_Sheet
5. IF any DELETE statement fails, THEN THE Add_On SHALL display the error message and continue processing remaining rows
6. IF the User cancels the fast delete dialog, THEN THE Add_On SHALL take no further action

### Requirement 10: Menu Registration

**User Story:** As a user, I want SQL and Table menus to appear when I open the spreadsheet, so that I can access all Add_On features from the menu bar.

#### Acceptance Criteria

1. WHEN the spreadsheet opens, THE Menu_Controller SHALL register an "SQL" menu with items: Show Prompt, Refresh, Clear History, Configure
2. WHEN the spreadsheet opens, THE Menu_Controller SHALL register a "Table" menu with items: Load Tables, Refresh Table, Drop Table, Fast Insert, Fast Update, Fast Delete
3. THE Menu_Controller SHALL use constant values for all menu labels instead of hard-coded strings

### Requirement 11: Error Handling and User Feedback

**User Story:** As a user, I want clear error messages when operations fail, so that I can understand what went wrong and take corrective action.

#### Acceptance Criteria

1. IF a database operation fails, THEN THE Add_On SHALL display an error dialog containing the error message
2. IF the User attempts a database operation without configured credentials, THEN THE Add_On SHALL display a message instructing the User to configure credentials first
3. WHEN a bulk operation processes multiple rows, THE Add_On SHALL report the count of successful and failed operations upon completion
4. IF a JDBC resource close operation fails, THEN THE JDBC_Service SHALL suppress the close error to avoid masking the original operation error

### Requirement 12: Result Rendering

**User Story:** As a user, I want query results displayed cleanly in the spreadsheet, so that I can read and work with the data easily.

#### Acceptance Criteria

1. WHEN a read query returns results, THE Result_Renderer SHALL write results starting at the current cursor position on the active sheet
2. THE Result_Renderer SHALL write one spreadsheet row per result row, with each column value in a separate cell
3. WHEN a read query returns zero data rows, THE Result_Renderer SHALL write only the status row indicating the query returned no results
4. THE Result_Renderer SHALL separate consecutive query result blocks with a blank row

### Requirement 13: Modular File Architecture

**User Story:** As a developer, I want the codebase organized into focused modules with clear responsibilities, so that the code is easier to maintain and test.

#### Acceptance Criteria

1. THE Add_On SHALL organize source code into separate files by responsibility: JDBC_Service, Config_Manager, Query_Builder, Input_Validator, Menu_Controller, Result_Renderer, Table operations, SQL prompt operations, and utility constants
2. THE Add_On SHALL define all sheet names and menu labels as constants in a single constants module
3. THE Add_On SHALL keep pure helper functions (value quoting, input validation, SQL classification) in modules that do not depend on Apps Script APIs
4. THE Add_On SHALL expose pure helper functions for testing with Node.js test scripts without requiring Apps Script runtime stubs

### Requirement 14: Testability

**User Story:** As a developer, I want comprehensive tests for all pure business logic, so that I can catch regressions without needing the Apps Script runtime.

#### Acceptance Criteria

1. THE Add_On SHALL provide Node.js-compatible test scripts for all pure helper functions
2. THE Add_On SHALL test the Query_Builder value quoting for strings, numbers, null values, empty strings, Date values, and strings containing single quotes
3. THE Add_On SHALL test the Input_Validator for empty input, whitespace-only input, and valid SQL statements
4. THE Add_On SHALL test SQL statement classification (identifying SELECT, SHOW, DESCRIBE vs write statements)
5. FOR ALL valid values, THE test suite SHALL verify that quoting a value and parsing the quoted result produces an equivalent value (round-trip property)
