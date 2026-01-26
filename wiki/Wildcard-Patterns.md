SSMS EnvTabs uses **SQL LIKE** syntax for matching.

## Syntax

| Symbol | Matches | Example |
|---|---|---|
| `%` | Any text (0+ chars) | `PROD%` matches `PROD`, `PROD-01` |
| `_` | Single character | `SQL_1` matches `SQL01`, `SQLA1` |

*Matching is typically case-insensitive.*

## Common Patterns

| Pattern | Goal | Example |
|---|---|---|
| `%` | Match **Everything** | Catch-all rule for leftovers. |
| `%PROD%` | Contains "PROD" | `SRV-PROD-01`, `MYPROD`. |
| `QA-%` | Starts with "QA-" | `QA-SQL01`. |
| `%-DB` | Ends with "-DB" | `Sales-DB`. |
| `SQL-__` | Specific Length | `SQL-01`, `SQL-02`. |
