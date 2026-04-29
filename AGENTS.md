# AGENTS.md

## Scope
These instructions apply to the entire repository.

## Project context
- This repository is a **Google Sheets Add-on / Apps Script** project.
- Runtime is Google Apps Script (V8). Favor plain JavaScript compatible with Apps Script.

## Coding guidelines
- Keep business logic in small helper functions.
- Prefer constants for sheet names/menu labels instead of hard-coded strings.
- Use strict comparisons (`===`, `!==`) unless coercion is intentionally required.
- Defensively handle cancel/no-op user actions from `Browser.inputBox`.
- Close JDBC resources (`ResultSet`, `Statement`, `Connection`) in `finally` blocks.
- Avoid dead commented-out code; delete obsolete blocks.

## Testing guidelines
- Add lightweight automated checks for pure helper logic whenever possible.
- Since Apps Script APIs are not available in CI, isolate testable pure functions and test with simple Node scripts.
- Include commands and outcomes in the final response.

## Documentation guidelines
- Keep README focused on add-on setup, permissions, supported DBs, and operational caveats.
