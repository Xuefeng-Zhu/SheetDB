## SheetDB

SheetDB connects Google Sheets to a remote SQL database so you can query data, edit rows, and sync table changes from a spreadsheet UI.

## Supported Databases

- Google Cloud SQL
- MySQL
- Microsoft SQL Server
- Oracle

## Setup

1. Open the spreadsheet and refresh the Apps Script project so custom menus load.
2. In the **SQL** menu, click **Configure**.
3. Provide:
   - JDBC URL (example: `jdbc:mysql://example.com:3306/db_name`)
   - Admin username
   - Admin password
4. After configuration completes, SheetDB can load and refresh tables.

> Notes:
> - Credentials are stored in the `Configue` sheet for backward compatibility with existing deployments.
> - Use least-privilege database credentials in production.

## Core Features

- **Show prompt**: Execute a SQL statement from a dialog prompt.
- **Refresh**: Re-run a SQL statement from the SQL history sheet.
- **Clear History**: Clear the active sheet history.
- **Load Tables**: Load all tables from the connected database into sheets.
- **Refresh Table**: Reload schema + rows for a selected table sheet.
- **Drop Table**: Drop a table in database and remove its sheet.
- **Fast Insert / Update / Delete**: Apply bulk row operations from selected ranges.

## Usage Tips

- Fast Update uses **bold cell formatting** to identify updated columns in selected rows.
- Operations run against the current table sheet automatically.
- If run from the `SQL` sheet, prompts ask for table name and attributes.

## Development Notes

- Source files are in `src/` and are intended for Google Apps Script runtime.
- The project uses helper utilities for:
  - configuration sheet creation / credential retrieval,
  - SQL sheet creation,
  - SQL-safe value quoting,
  - and centralized JDBC execution error handling.

## Related Project

[SheetSQL](https://github.com/Xuefeng-Zhu/SheetSQL) processes data with SQL directly in Google Sheets.

## License

MIT
