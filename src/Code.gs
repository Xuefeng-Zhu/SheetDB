function onOpen() {
  var ss = SpreadsheetApp.getActive();
  var items = [
    {name: 'Show prompt', functionName: 'showPrompt'},
    {name: 'Refresh', functionName: 'refresh'},
    {name: 'Clear History', functionName: 'warning'},
    {name: 'Configue', functionName: 'configue'}
  ];
  var table = [
    {name: 'Load Tables', functionName: 'loadTables'},
    {name: 'Refresh Table', functionName: 'refreshTable'},
    {name: 'Drop Table', functionName: 'dropTable'},
    {name: 'Fast Insert', functionName: 'fastInsert'},
    {name: 'Fast Update', functionName: 'fastUpdate'},
    {name: 'Fast Delete', functionName: 'fastDelete'}

  ];
  ss.addMenu('SQL', items);
  ss.addMenu('Table', table);
}

function showPrompt() {
  var result = Browser.inputBox(
      'Google Sheet based SQL',
      'Please enter SQL statement you want to execute:',
      Browser.Buttons.OK_CANCEL);

  if (result != 'cancel') {
    var out = SQL(result);
    var sheet = SpreadsheetApp.getActiveSheet();
    for (var i = 0; i < out.length; i++)
      sheet.appendRow(out[i]);
    sheet.appendRow([" "]);

  } else {
    Browser.msgBox('Thanks for using! Bye!');
  }
}

function refresh(){
  var cell = SpreadsheetApp.getActiveRange();
  var statement = cell.getValue();
  statement = statement.replace("Success","").trim();
  var out = SQL(statement);
  var sheet = cell.getSheet();
  var temp = cell.getRowIndex();
  var nrow = 1;
  while (sheet.getRange(temp + nrow, 1).getValue() != " " && (!sheet.getRange(temp + nrow, 1).isBlank()))
    nrow++;
  if ( nrow < out.length)
     sheet.insertRows(temp, out.length - nrow);
  else if (nrow > out.length)
    sheet.deleteRows(temp, nrow - out.length);
  
  sheet.deleteRow(temp);
  sheet.insertRows(temp);
  sheet.getRange(temp, 1).setValue(out[0])
  out.shift();
  var range = sheet.getRange(temp + 1, 1, out.length, out[0].length);
  range.setValues(out);
}

function warning(){
  var result = Browser.msgBox(
    'Please confirm',
    'Are you sure you want to clear all the history?',
    Browser.Buttons.YES_NO);

  if (result == 'yes') {
    var sheet = SpreadsheetApp.getActiveSheet();
    sheet.clear();
    Browser.msgBox('History Cleared.');
  } else {
    Browser.msgBox('User Canceled.');
  }
}

function configue(){
  var url = Browser.inputBox(
      'Configue',
      'Please enter the URL of SQL database:',
      Browser.Buttons.OK_CANCEL);    
  if (url == 'cancel') {
    Browser.msgBox('Configuration does not complete!');
    return;
  }
  
  var adm = Browser.inputBox(
    'Configue',
    'Please enter the administrator of SQL database:',
    Browser.Buttons.OK_CANCEL);  
  if (adm == 'cancel'){
    Browser.msgBox('Configuration does not complete!');
    return;
  }
  
  var password = Browser.inputBox(
    'Configue',
    'Please enter the password of the administrator:',
    Browser.Buttons.OK_CANCEL);  
  if (password == 'cancel') {
    Browser.msgBox('Configuration does not complete!');
    return;
  }
  
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName("Configue");
  sheet.getRange(1, 2).setValue(url);
  sheet.getRange(2, 2).setValue(adm);
  sheet.getRange(3, 2).setValue(password);
  Browser.msgBox('Configuration completes!');
  loadTables();
}

function SQL(input) {
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
  out.push([input + " Success"]);
  var temp = input.split(" ");
  if (temp[0].toUpperCase() == "SELECT" || temp[0].toUpperCase() == "SHOW" ||  temp[0].toUpperCase() == "DESCRIBE")
  {
    result = statement.executeQuery(input);
    while (result.next())
    {
      var temp = [];
      for (var col = 0; col < result.getMetaData().getColumnCount(); col++) 
        temp.push(result.getString(col + 1));
      out.push(temp);
    }
  }
  else 
  {
    result = statement.executeUpdate(input);
  }

  return out;
}

