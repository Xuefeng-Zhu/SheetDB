##Overview
The SQL implementation is based on Google sheet by connecting to remote MySQL database system via JDBC.

##Configue
Before starting to use the tool, it need to be configued by clicking Configure in SQL menu. After the url, adm, and password are stored, you are able to run SQL statement in Google Sheet.
There are two ways to run the SQL statement. You can type directly in cell, like =SQL() to run it, or click SQL on the Menu.

##Feature
There are several advanced features with this tool:
```
1. Add Table option into menu, which contains functions like Load Tables, Refresh Table, Drop Table, Fast Insert, Fast Update,
Fast Delete
2. After the database is configured, the system will automatically load tables in database into google sheet.  Similar to
Approach2, the table will have attributes and values.  Later user can use load tables to load new tables or refresh tables.
3. Refresh table is used to refresh a single table
4. Drop Table will drop the table in both google sheet and database 
5. Fast Insert is used to insert a range of data. The way to use it is select a range of data for inserting and click the Fast
Insert function in Table menu.
6. Fast Update is used to update a range of data. The way to use it is to make chances on data, the cell will be automatically
marked as bold, then select the rows which is made changes and click Fast Update function in Table menu.
7. Fast Delete is used to delete a range of data. The way to use it is select a range of data for inserting and click the Fast
Delete function in Table menu.
8. To refresh SQL statement in history. The way to use it is to select the statement you want to reexecute in the history sheet,
and then click Refresh in SQL menu.
```
All the functions in Table menu will automatically apply to current table sheet, but if user uses them in SQL Sheet, table name and attributes of table will be asked as an input. In SQL sheet, user can use Fast Insert, Fast Update, and Fast Delete to quickly operate data returned by SQL Query.

I also tried to implement the event trigger function to allow data in database to be automatically updated when user makes change to Google sheet, but after testing it, I found JDBC is not working in the trigger function. I guess the reason is that trigger function is not supporting external library like JDBC, which will cost lots of computation time. 


The link to the project is [demo](https://docs.google.com/spreadsheet/ccc?key=0AlMMHFOg-bRZdHlLR0hEZDBQakhQQ2NsdkJ2NGwyeVE&usp=sharing)
