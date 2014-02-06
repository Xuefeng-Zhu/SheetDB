/*
function onEdit(event) {
  var sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getSheetName() != "SQL")
  {
    var name = sheet.getSheetName();
    var attributes = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var value = event.value;
    var rowI = event.range.getRowIndex();
    var colI = event.range.getColumn() - 1;
    var row = sheet.getRange(rowI, 1, 1, attributes.length).getValues()[0];
    var update = null;
    if (typeof event.range.getValue() == "string")
      update = attributes[colI] + "=" + "'"+ event.value + "'";
    else
      update = attributes[colI] + "=" + event.value;
    
    var condition = "";
    for (var c = 0; c < attributes.length; c++)
    {
      if (c == colI) 
        continue;
      if (typeof row[c] == "string")
        row[c] = "'"+ row[c] + "'";

      if (row[c] === +row[c] && row[c] !== (row[c]|0))
        continue;
      
      if (condition.length > 0)
        condition += " AND ";
      condition += (attributes[c] + "=" + row[c]);
    }

    var input = "UPDATE " + name + " SET " + update + " WHERE " + condition;
    var ss = SpreadsheetApp.getActive();
    sheet = ss.getSheetByName("Configue");
    var url = sheet.getRange(1, 2).getValue(); 
    var adm = sheet.getRange(2, 2).getValue(); 
    var password = sheet.getRange(3, 2).getValue(); 
    if (url == "" || adm == "" || password == ""){
      Browser.msgBox("Please configue the system first!");
      return;
    }

    Jdbc.getConnection(url, adm, password);
                Browser.msgBox(password);

    var statement = conn.createStatement();

    statement.executeUpdate(input);
    sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SQL");
    sheet.appendRow([input + " Success"]);
  }
}
*/

function onEdit(event)
{
  event.range.setFontWeight("bold");
}