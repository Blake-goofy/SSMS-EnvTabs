# Wildcard Patterns

Guide to using wildcard patterns for server and database matching in SSMS EnvTabs.

## Pattern Syntax

SSMS EnvTabs uses **SQL LIKE** syntax for pattern matching, not regular expressions.

### Available Wildcards

| Wildcard | Matches | Example Pattern | Matches | Doesn't Match |
|----------|---------|-----------------|---------|---------------|
| `%` | Any sequence of characters (0 or more) | `PROD%` | PROD, PROD-01, PRODUCTION | PREPROD |
| `_` | Exactly one character | `SQL-0_` | SQL-01, SQL-02, SQL-09 | SQL-0, SQL-010 |
| Exact text | Exact match (no wildcards) | `SERVER-01` | SERVER-01 only | SERVER-02, SERVER-01-A |

**Important**: 
- `%` matches **zero or more** characters, including empty string
- `_` matches **exactly one** character
- Matching is typically **case-insensitive**

## Common Patterns

### Exact Match

Match a specific server or database name:

```json
{
  "server": "PROD-SQL-01",
  "database": "CustomerDB"
}
```

**Matches**:
- Server: `PROD-SQL-01` AND Database: `CustomerDB`

**Doesn't match**:
- `PROD-SQL-02` (different server)
- Server `PROD-SQL-01` with database `OrdersDB` (different database)

### Match All

Match any server or any database:

```json
{
  "server": "%",
  "database": "%"
}
```

**Matches**: Everything

**Use case**: Catch-all rule at the end of your rule list

### Contains Pattern

Match servers/databases containing specific text:

```json
{
  "server": "%PROD%",
  "database": "%"
}
```

**Matches servers**:
- `PROD-SQL-01`
- `SERVER-PROD-02`
- `MY-PRODUCTION-DB`
- `PREPROD` (contains "PROD")

**Doesn't match**:
- `DEV-SQL-01`
- `QA-SERVER`

### Starts With

Match servers/databases starting with specific text:

```json
{
  "server": "PROD-%",
  "database": "Customer%"
}
```

**Matches**:
- Server: `PROD-SQL-01`, `PROD-APP-05`
- Database: `CustomerDB`, `CustomerOrders`, `Customers`

**Doesn't match**:
- Server: `MY-PROD-SERVER` (doesn't start with "PROD-")
- Database: `AllCustomers` (doesn't start with "Customer")

### Ends With

Match servers/databases ending with specific text:

```json
{
  "server": "%-PROD",
  "database": "%DB"
}
```

**Matches**:
- Server: `SERVER-PROD`, `SQL-01-PROD`
- Database: `CustomerDB`, `OrdersDB`

**Doesn't match**:
- Server: `PROD-SERVER` (doesn't end with "-PROD")
- Database: `Database1` (doesn't end with "DB")

### Single Character Wildcard

Match patterns with specific length/format:

```json
{
  "server": "SQL-0_",
  "database": "DB__"
}
```

**Matches**:
- Server: `SQL-01`, `SQL-02`, ..., `SQL-09`
- Database: `DB01`, `DB02`, `DBAA`

**Doesn't match**:
- Server: `SQL-0` (too short), `SQL-010` (too long)
- Database: `DB1` (too short), `DB001` (too long)

## Real-World Examples

### Example 1: Multiple Production Servers

```json
{
  "groupName": "Prod",
  "server": "%PROD%",
  "database": "%",
  "priority": 10,
  "colorIndex": 3
}
```

**Matches**:
- `PROD-SQL-01`
- `PROD-SQL-02`
- `MY-PROD-SERVER`
- `SERVER-PRODUCTION-01`
- `PREPROD` (be careful! might want to exclude)

### Example 2: Specific Database Across All Servers

```json
{
  "groupName": "Finance",
  "server": "%",
  "database": "FinanceDB",
  "priority": 20,
  "colorIndex": 11
}
```

**Matches**: Any server + `FinanceDB` database

### Example 3: Test Databases

```json
{
  "groupName": "Test",
  "server": "%",
  "database": "%Test%",
  "priority": 50,
  "colorIndex": 9
}
```

**Matches databases**:
- `TestDB`
- `CustomerTest`
- `Test_Orders`
- `PerformanceTestResults`

### Example 4: Regional Servers

```json
{
  "groups": [
    {
      "groupName": "US-Prod",
      "server": "US-%-PROD",
      "database": "%",
      "priority": 10,
      "colorIndex": 3
    },
    {
      "groupName": "EU-Prod",
      "server": "EU-%-PROD",
      "database": "%",
      "priority": 11,
      "colorIndex": 3
    }
  ]
}
```

**Matches**:
- US-Prod: `US-EAST-PROD`, `US-WEST-PROD`, `US-CENTRAL-PROD`
- EU-Prod: `EU-NORTH-PROD`, `EU-WEST-PROD`

### Example 5: Environment Prefixes

```json
{
  "groups": [
    {
      "groupName": "Prod",
      "server": "PROD-%",
      "database": "%",
      "priority": 10,
      "colorIndex": 3
    },
    {
      "groupName": "QA",
      "server": "QA-%",
      "database": "%",
      "priority": 20,
      "colorIndex": 1
    },
    {
      "groupName": "Dev",
      "server": "DEV-%",
      "database": "%",
      "priority": 30,
      "colorIndex": 4
    }
  ]
}
```

**Matches**:
- Prod: `PROD-SQL-01`, `PROD-APP-02`
- QA: `QA-SQL-01`, `QA-TEST-SERVER`
- Dev: `DEV-WORKSTATION`, `DEV-SQL-LOCAL`

### Example 6: Excluding Specific Patterns

Use priority to match specific patterns before broader ones:

```json
{
  "groups": [
    {
      "groupName": "QA",
      "server": "%QA%",
      "database": "%",
      "priority": 20,
      "colorIndex": 1,
      "comment": "Match QA servers..."
    },
    {
      "groupName": "Prod",
      "server": "%PROD%",
      "database": "%",
      "priority": 10,
      "colorIndex": 3,
      "comment": "...but check PROD first (lower priority number)"
    }
  ]
}
```

**Problem scenario**: Server named `PROD-QA-01` could match both rules

**Solution**: The rule with priority 10 (Prod) is checked first and matches, so it becomes "Prod-1", not "QA-1"

## Advanced Patterns

### Combining Multiple Wildcards

```json
{
  "server": "%-%-%",
  "database": "DB__%"
}
```

**Matches**:
- Server: Any server with at least 2 hyphens (e.g., `US-PROD-01`, `SERVER-APP-DEV`)
- Database: Databases starting with "DB" and at least 2 more characters (e.g., `DB01`, `DBAA`)

### Specific Format Matching

```json
{
  "server": "SQL-____-__",
  "database": "%"
}
```

**Matches servers**: `SQL-PROD-01`, `SQL-TEST-99`
**Format**: "SQL-" + 4 characters + "-" + 2 characters

### Negative Patterns (Using Priority)

You can't directly exclude patterns, but use priority to match desired patterns first:

```json
{
  "groups": [
    {
      "groupName": "Prod-Critical",
      "server": "PROD-CORE-%",
      "database": "%",
      "priority": 5,
      "colorIndex": 3,
      "comment": "Most specific - matches first"
    },
    {
      "groupName": "Prod",
      "server": "PROD-%",
      "database": "%",
      "priority": 10,
      "colorIndex": 7,
      "comment": "Broader - matches remaining PROD servers"
    }
  ]
}
```

## Pattern Matching Tips

### Tip 1: Order Matters (Priority)

Rules are evaluated by priority (lowest first). **First match wins.**

```json
// WRONG ORDER
{
  "groups": [
    {
      "server": "%",              // Matches everything
      "priority": 10
    },
    {
      "server": "%PROD%",        // Never reached!
      "priority": 20
    }
  ]
}

// CORRECT ORDER
{
  "groups": [
    {
      "server": "%PROD%",        // Specific first
      "priority": 10
    },
    {
      "server": "%",              // Catch-all last
      "priority": 100
    }
  ]
}
```

### Tip 2: Test Your Patterns

Test patterns with a temporary high-priority rule:

```json
{
  "groupName": "DEBUG",
  "server": "%",
  "database": "%",
  "priority": 1,
  "colorIndex": 15
}
```

All tabs will become "DEBUG-1", "DEBUG-2", etc. Remove after testing.

### Tip 3: Be Specific When Possible

```json
// LESS SPECIFIC
"server": "%PROD%"              // Matches PREPROD, PROD, PRODUCTION

// MORE SPECIFIC
"server": "PROD-%"              // Matches only PROD-xxx servers
```

### Tip 4: Document Complex Patterns

```json
{
  "groupName": "Prod-US-West",
  "server": "US-W_-PROD-__",
  "database": "%",
  "priority": 5,
  "colorIndex": 3,
  "comment": "Matches US-W1-PROD-01 through US-W9-PROD-99"
}
```

### Tip 5: Case Insensitivity

Most SQL Server installations are case-insensitive for server names:

```json
// These are typically equivalent
"server": "%PROD%"
"server": "%prod%"
"server": "%Prod%"
```

But to be safe, match the case of your actual server names.

## Common Mistakes

### Mistake 1: Using Wrong Wildcard Syntax

```json
// WRONG - These are NOT SQL LIKE wildcards
"server": "PROD*"          // Shell glob
"server": "PROD.*"         // Regex
"server": "PROD.+"         // Regex
"server": "PROD[0-9]"      // Regex

// CORRECT - SQL LIKE wildcards
"server": "PROD%"          // SQL LIKE
"server": "PROD_"          // SQL LIKE
```

### Mistake 2: Forgetting `%` at Start or End

```json
// WRONG - Matches only exact "PROD"
"server": "PROD"

// CORRECT - Matches anything containing "PROD"
"server": "%PROD%"
```

### Mistake 3: Too Broad Patterns

```json
// TOO BROAD - Matches PREPROD, PROD, PRODUCTION, etc.
"server": "%PROD%"

// BETTER - More specific
"server": "PROD-%"         // Only servers starting with PROD-
```

### Mistake 4: Not Testing Edge Cases

Always consider:
- What if server name has spaces? `"% PROD %"`
- What if server name has hyphens, underscores? `"%PROD_%" or "%PROD-%"`
- What about similar names? `PROD` vs `PREPROD`

## Testing Your Patterns

### Method 1: Use a Test Rule

Add a test rule at the top (priority 1):

```json
{
  "groupName": "TEST",
  "server": "YOUR-PATTERN-HERE",
  "database": "%",
  "priority": 1,
  "colorIndex": 15
}
```

If tabs are renamed to "TEST-1", your pattern matches!

### Method 2: Use Catch-All Rules

Add a catch-all at the end to see what's not matching:

```json
{
  "groupName": "Unmatched",
  "server": "%",
  "database": "%",
  "priority": 999,
  "colorIndex": 8
}
```

Any tabs that become "Unmatched-X" didn't match earlier rules.

### Method 3: Check Logs

Look at the log file to see what patterns are matching:

```
%LocalAppData%\SSMS EnvTabs\runtime.log
```

## Next Steps

- [Configuration Guide](Configuration-Guide.md) - Complete configuration reference
- [Tips & Best Practices](Tips-and-Best-Practices.md) - Optimize your patterns
- [Troubleshooting](Troubleshooting.md) - Pattern matching issues
