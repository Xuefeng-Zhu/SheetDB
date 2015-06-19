function loadTables()
{
  var tables = SQLHelp("SHOW tables");
  for (var i = 0; i < tables.length; i++)
  {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(tables[i][0]);
    if (sheet != null)
    {
      refreshTHelp(sheet);
      continue;
    }
    sheet = ss.insertSheet(tables[i][0], ss.getNumSheets());
    var temp = SQLHelp("DESCRIBE " + tables[i][0]);
    var attributes = [];
    for (var t = 0; t < temp.length; t++)
      attributes.push(temp[t][0]);
    sheet.appendRow(attributes);
    var cells = sheet.getRange(1, 1, 1, attributes.length);
    cells.setFontWeight("bold");
    temp = SQLHelp("SELECT * FROM " + tables[i][0]);
    if (temp.length > 0)
    {
      cells = sheet.getRange(2, 1, temp.length, temp[0].length);
      cells.setValues(temp);
    }
  }
}

function refreshTable()
{
  var sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getSheetName() == "SQL")
  {
    var result = Browser.inputBox(
      'Refresh Table',
      'Please enter table name for refresh:',
      Browser.Buttons.OK_CANCEL);
    if (result != 'cancel')
      sheet = SpreadsheetApp.getActive().getSheetByName(result);
    else
    {
      Browser.msgBox('Thanks for using! Bye!');
      return;
    }
  }
  refreshTHelp(sheet);
  Browser.msgBox('Table successfully refreshed');
}

function refreshTHelp(sheet)
{
  var tableName = sheet.getSheetName();
  sheet.clear();
  var temp = SQLHelp("DESCRIBE " + tableName);
  var attributes = [];
  for (var t = 0; t < temp.length; t++)
    attributes.push(temp[t][0]);
  sheet.appendRow(attributes);
  var cells = sheet.getRange(1, 1, 1, attributes.length);
  cells.setFontWeight("bold");
  temp = SQLHelp("SELECT * FROM " + tableName);
  if (temp.length > 0)
  {
    cells = sheet.getRange(2, 1, temp.length, temp[0].length);
    cells.setValues(temp);
  }
}

function dropTable()
{
  var sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getSheetName() == "SQL")
  {
    var result = Browser.inputBox(
      'Drop Table',
      'Please enter table name for drop:',
      Browser.Buttons.OK_CANCEL);
    if (result != 'cancel')
      sheet = SpreadsheetApp.getActive().getSheetByName(result);
    else
    {
      Browser.msgBox('Thanks for using! Bye!');
      return;
    }
  }
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SQL").appendRow(SQL("DROP TABLE " + sheet.getSheetName())[0]);
  SpreadsheetApp.getActiveSpreadsheet().deleteSheet(sheet);
}

function fastInsert()
{
  var result = null;
  var sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getSheetName() == "SQL")
  {
    var result = Browser.inputBox(
      'Fast Insert',
      'Please enter table name (and attributes of inserted data seperate by  comma):',
      Browser.Buttons.OK_CANCEL);
  }
  else 
  {
    var name = sheet.getSheetName();
    result = name;
  }
  
  if (result != 'cancel') {
    var temp = result.split(",");
    var tableName = temp[0];
    var attributes = "";
    if (temp.length > 1)
    {
      attributes = temp.slice(1).join();
      attributes = " (" + attributes + ")";
    }
    var range = SpreadsheetApp.getActiveRange();
    var values = range.getValues();
    for (var r = 0; r < values.length; r++)
    {
      var row = values[r];
      for (var c = 0; c < row.length; c++)
      {
        if (typeof row[c] == "string")
          row[c] = "'"+ row[c] + "'";
      }
      row = row.join(",");
      row = "(" + row + ")";
      var out = SQL("INSERT INTO " + tableName + attributes + " VALUES " + row);
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SQL");
      for (var i = 0; i < out.length; i++)
        sheet.appendRow(out[i]);
    }
    sheet.appendRow([" "]);
    sheet = SpreadsheetApp.getActiveSheet();
    if (sheet.getSheetName() != "SQL")
      range.clearFormat();

  } else {
    Browser.msgBox('Thanks for using! Bye!');
  }
}

function fastDelete()
{
  var result = null;
  var sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getSheetName() == "SQL")
  {
     result = Browser.inputBox(
        'Fast Delete',
        'Please enter table name and table attributes seperate by comma:',
        Browser.Buttons.OK_CANCEL);
  }
  else 
  {
    var name = sheet.getSheetName();
    var attributes = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    result = name + "," + attributes.join();
  }
  
  if (result != 'cancel') {
    var temp = result.split(",");
    var tableName = temp[0];
    var attributes = temp.slice(1);
    var range = SpreadsheetApp.getActiveRange();
    var values = range.getValues();
    for (var r = 0; r < values.length; r++)
    {
      var row = values[r];
      var temp = "";
      for (var c = 0; c < row.length; c++)
      {
        if (row[c] === +row[c] && row[c] !== (row[c]|0))
          continue;
        if (c != 0 && temp != "")
          temp += " AND ";
        if (typeof row[c] == "string")
          row[c] = "'"+ row[c] + "'";
        temp += (attributes[c] + "=" + row[c]);
      }
      var out = SQL("DELETE FROM " + tableName + " WHERE " + temp);
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SQL");
      for (var i = 0; i < out.length; i++)
        sheet.appendRow(out[i]);
    }
    sheet.appendRow([" "]);
    
    sheet = SpreadsheetApp.getActiveSheet();
    if (sheet.getSheetName() != "SQL")
    {
      var rowI = range.getRowIndex();
      var n = range.getNumRows();
      sheet.deleteRows(rowI, n);
    }

  } 
  else 
  {
    Browser.msgBox('Thanks for using! Bye!');
  }
}

function fastUpdate()
{
  
  var result = null;
  var sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getSheetName() == "SQL")
  {
    var result = Browser.inputBox(
      'Fast Update',
      'Please enter table name and table attributes seperate by comma:',
      Browser.Buttons.OK_CANCEL);
  }
  else 
  {
    var name = sheet.getSheetName();
    var attributes = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    result = name + "," + attributes.join();
  }
  
  if (result != 'cancel') {
    var temp = result.split(",");
    var tableName = temp[0];
    var attributes = temp.slice(1);
    var range = SpreadsheetApp.getActiveRange();
    var values = range.getValues();
    var weights = range.getFontWeights();
    for (var r = 0; r < values.length; r++)
    {
      var row = values[r];
      var weight = weights[r];
      var update = "";
      var condition = "";
      for (var c = 0; c < row.length; c++)
      {
        if (typeof row[c] == "string")
          row[c] = "'"+ row[c] + "'";

        if (weight[c] == "bold")
        {
          if (update.length > 0)
            update += " , ";
          update += (attributes[c] + "=" + row[c]);
        }
        else 
        {
          if (row[c] === +row[c] && row[c] !== (row[c]|0))
            continue;
          if (condition.length > 0)
            condition += " AND ";
          condition += (attributes[c] + "=" + row[c]);
        }
      }
      if (update.length > 0)
      {
        var out = SQL("UPDATE " + tableName + " SET " + update + " WHERE " + condition);
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SQL");
        for (var i = 0; i < out.length; i++)
          sheet.appendRow(out[i]);
      }
    }
    sheet.appendRow([" "]);
    sheet = SpreadsheetApp.getActiveSheet();
    if (sheet.getSheetName() != "SQL")
      range.clearFormat();
  } else {
    Browser.msgBox('Thanks for using! Bye!');
  }
}

function SQLHelp(input) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName("Configue");
  var url = sheet.getRange(1, 2).getValue(); 
  var adm = sheet.getRange(2, 2).getValue(); 
  var password = sheet.getRange(3, 2).getValue(); 
  
  if (url == "" || adm == "" || password == ""){
    Browser.msgBox("Please configue the system first!");
    return;
  }
      
  var conn = Jdbc.getConnection(url, adm, password);
  var statement = conn.createStatement();
  var result;
  var out = new Array();
  result = statement.executeQuery(input);
  while (result.next())
  {
    var temp = [];
    for (var col = 0; col < result.getMetaData().getColumnCount(); col++) 
      temp.push(result.getString(col + 1));
    out.push(temp);
  }
  return out;
}

