/**
 * MenuController — registers menus on spreadsheet open.
 *
 * Depends on: SpreadsheetApp, Constants (menu label constants),
 *             SqlPrompt (showPrompt, refresh, clearHistory),
 *             TableOps (loadTables, refreshTable, dropTable, fastInsert, fastUpdate, fastDelete),
 *             ConfigManager (runConfigDialog)
 */

/**
 * Registers the SQL and Table menus when the spreadsheet is opened.
 * Routes each menu item to the appropriate handler function.
 */
function onOpen() {
  var ss = SpreadsheetApp.getActive();

  var sqlItems = [
    {name: MENU_SHOW_PROMPT, functionName: 'showPrompt'},
    {name: MENU_REFRESH, functionName: 'refresh'},
    {name: MENU_CLEAR_HISTORY, functionName: 'clearHistory'},
    {name: MENU_CONFIGURE, functionName: 'runConfigDialog'}
  ];

  var tableItems = [
    {name: MENU_LOAD_TABLES, functionName: 'loadTables'},
    {name: MENU_REFRESH_TABLE, functionName: 'refreshTable'},
    {name: MENU_DROP_TABLE, functionName: 'dropTable'},
    {name: MENU_FAST_INSERT, functionName: 'fastInsert'},
    {name: MENU_FAST_UPDATE, functionName: 'fastUpdate'},
    {name: MENU_FAST_DELETE, functionName: 'fastDelete'}
  ];

  ss.addMenu(MENU_SQL, sqlItems);
  ss.addMenu(MENU_TABLE, tableItems);
}

/**
 * Backward-compatible alias for runConfigDialog.
 * Matches the old Code.gs "configue()" function name.
 */
function configue() {
  runConfigDialog();
}
