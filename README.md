## SheetDB
SheetDB allows you to connect to remote SQL server and load data into Google Sheet directly. Porting data from database to Google Sheet allows you easily modify and visualize data.

[demo](https://docs.google.com/spreadsheet/ccc?key=0AlMMHFOg-bRZdHlLR0hEZDBQakhQQ2NsdkJ2NGwyeVE&usp=sharing)

**Related Project**: 

[SheetSQL](https://github.com/Xuefeng-Zhu/SheetSQL) allows you process data using SQL directly in Google Sheet

## Usage
Before starting to use SheetDC, it needs to be configued. Click Configure in SQL menu and enter url, admin, and password. **Currently, SheetDB only supports  Google Cloud SQL, MySQL, Microsoft SQL Server, and Oracle databases.** If everything sets up correctly, you are able to run SQL statement in Google Sheet.

There are two ways to run SQL statements. You can type directly in cell, like `=SQL()` to run it, or click SQL on the Menu.

## Features

* After the SheetDB is configured, SheetDB will automatically load tables in database into google sheet. Later user can use Load Tables to load new tables or Refresh Tables to fetch latest data.
* In the menu, Table option contains functionalities like Load Tables, Refresh Table, Drop Table, Fast Insert, Fast Update,
Fast Delete
	* Refresh table will refresh the table you selected
	* Drop Table will drop the table you selected in both Google sheet and database 
	* Fast Insert will insert a range of data into remote database. Select a range of data for inserting and click the Fast Insert. If you are at SQL sheet, you will be required to enter table name and attribute 
	* Fast Update will update a range of data to remote database. Make some changes on data, and cells will be marked as bold if they got modified. Select the rows which are modified and click Fast Update.
	* Fast Delete will delete a range of data in remote database. Select a range of data and click the Fast Delete.
8. To refresh SQL statement in history, select the statement you want to reexecute in the SQL sheet,
and then click Refresh in SQL menu.

All the functions in Table menu will automatically apply to current table sheet, but if user uses them in SQL Sheet, table name and attributes of table will be asked as an input. In SQL sheet, user can use Fast Insert, Fast Update, and Fast Delete to quickly operate data returned by SQL Query.

## Licence
MIT


