# Configuration Guide

Complete reference for configuring SSMS EnvTabs.

## Configuration File Location

The configuration is stored in a JSON file at:

```
%USERPROFILE%\Documents\SSMS EnvTabs\TabGroupConfig.json
```

For example: `C:\Users\YourName\Documents\SSMS EnvTabs\TabGroupConfig.json`

## Opening the Configuration File

### From SSMS (Recommended)

1. Open SSMS
2. Go to **Tools** → **SSMS EnvTabs** → **Open Configuration**
3. The configuration file opens in SSMS for editing
4. Make your changes and save (Ctrl+S)
5. Changes take effect immediately

### Manually

1. Navigate to `%USERPROFILE%\Documents\SSMS EnvTabs\` in File Explorer
2. Open `TabGroupConfig.json` in any text editor
3. Make your changes and save
4. Changes take effect when you open/switch tabs

## Configuration Structure

The configuration file has three main sections:

```json
{
  "description": "Configuration for SSMS EnvTabs",
  "version": "1.2.0",
  "settings": {
    // Global settings
  },
  "documentation": {
    // Inline help text (optional)
  },
  "colors": {
    // Color reference (optional)
  },
  "groups": [
    // Array of tab group rules
  ]
}
```

## Configuration Options

### Settings Section

The `settings` section controls global behavior:

```json
"settings": {
  "enableAutoRename": true,
  "enableAutoColor": true,
  "autoConfigure": "server db",
  "enableConfigurePrompt": true
}
```

#### enableAutoRename

- **Type**: Boolean (`true` or `false`)
- **Default**: `true`
- **Description**: Enables automatic tab renaming

**Behavior**:
- When `true`: Only **unsaved** query tabs (e.g., "SQLQuery1.sql") are renamed
- Saved files keep their original names to preserve your work
- Renamed tabs show the group name + counter (e.g., "Prod-1", "Prod-2")

**Example**:
```json
"enableAutoRename": true
```

#### enableAutoColor

- **Type**: Boolean (`true` or `false`)
- **Default**: `false`
- **Description**: Enables automatic tab coloring

**Behavior**:
- When `true`: Tabs are colored based on the matched rule's `colorIndex`
- When `false`: Tabs use default SSMS coloring
- Requires SSMS to support tab coloring (SSMS 22.0+)

**Example**:
```json
"enableAutoColor": true
```

#### autoConfigure

- **Type**: String
- **Values**: `"off"`, `"server"`, `"server db"`
- **Default**: `"off"`
- **Description**: Controls automatic rule creation for new connections

**Options**:

1. **`"off"`** - Disabled
   - No automatic rules are created
   - You must manually configure all rules

2. **`"server"`** - Group by server only
   - Creates one rule per unique server
   - All databases on that server share the same group
   - Example: "SERVER-01" gets a rule matching any database

3. **`"server db"`** - Group by server and database
   - Creates one rule per unique server+database combination
   - Most granular auto-configuration
   - Example: "SERVER-01:CustomerDB" and "SERVER-01:OrdersDB" get separate rules

**Example**:
```json
"autoConfigure": "server db"
```

#### enableConfigurePrompt

- **Type**: Boolean (`true` or `false`)
- **Default**: `true`
- **Description**: Shows a prompt after auto-creating rules

**Behavior**:
- When `true`: After auto-creating a rule, asks if you want to edit the configuration
- When `false`: Silently creates rules without prompting
- Only applies when `autoConfigure` is not `"off"`

**Example**:
```json
"enableConfigurePrompt": true
```

## Group Rules

The `groups` array contains rules that define how tabs are matched, renamed, and colored.

### Basic Rule Structure

```json
{
  "groupName": "Prod",
  "server": "%PROD%",
  "database": "%",
  "priority": 10,
  "colorIndex": 3
}
```

### Rule Fields

#### groupName

- **Type**: String
- **Required**: Yes
- **Description**: The prefix shown on matching tabs

**Guidelines**:
- Keep it short (4-8 characters work best)
- Should be distinctive and recognizable
- Used with a counter: "Prod-1", "Prod-2", etc.

**Examples**:
```json
"groupName": "Prod"      // Results in: Prod-1, Prod-2, Prod-3...
"groupName": "QA"        // Results in: QA-1, QA-2, QA-3...
"groupName": "Dev-Local" // Results in: Dev-Local-1, Dev-Local-2...
```

#### server

- **Type**: String
- **Required**: Yes
- **Description**: Server name pattern using SQL LIKE syntax

**Wildcard Support**:
- `%` - Matches any sequence of characters (including zero characters)
- `_` - Matches any single character

**Examples**:
```json
"server": "PROD-SQL-01"        // Exact match only
"server": "%PROD%"             // Matches any server with "PROD" in the name
"server": "PROD-%"             // Matches servers starting with "PROD-"
"server": "%-PROD-%"           // Matches servers with "-PROD-" anywhere
"server": "%"                  // Matches all servers
"server": "SQL-0_"             // Matches SQL-01, SQL-02, ..., SQL-09
```

**Case Sensitivity**: Server matching is typically case-insensitive

#### database

- **Type**: String
- **Required**: Yes
- **Description**: Database name pattern using SQL LIKE syntax

Uses the same wildcard syntax as `server` field.

**Examples**:
```json
"database": "CustomerDB"       // Exact match only
"database": "%Test%"           // Any database with "Test" in the name
"database": "App_%"            // Databases starting with "App_"
"database": "%"                // Matches all databases
```

#### priority

- **Type**: Integer
- **Required**: Yes
- **Description**: Determines the order rules are evaluated (lower = first)

**How Priority Works**:
1. Rules are sorted by priority (lowest number first)
2. Rules are evaluated in order
3. **First matching rule wins** - evaluation stops
4. Lower-priority rules are never checked if a higher-priority rule matches

**Best Practices**:
- Use increments of 10 (10, 20, 30...) to allow inserting rules later
- Specific rules should have lower numbers (10, 20)
- General/fallback rules should have higher numbers (90, 100)

**Example Priority Strategy**:
```json
// Priority 10 - Most specific
{
  "groupName": "Prod-Critical",
  "server": "PROD-SQL-01",
  "database": "CustomerDB",
  "priority": 10
}

// Priority 20 - Less specific
{
  "groupName": "Prod",
  "server": "%PROD%",
  "database": "%",
  "priority": 20
}

// Priority 30 - General
{
  "groupName": "QA",
  "server": "%QA%",
  "database": "%",
  "priority": 30
}

// Priority 100 - Catch-all
{
  "groupName": "Other",
  "server": "%",
  "database": "%",
  "priority": 100
}
```

#### colorIndex

- **Type**: Integer (0-15)
- **Required**: Yes
- **Description**: The tab color index

**Values**: Must be between 0 and 15 (inclusive)

See the [Color Reference](Color-Reference.md) for all available colors with visual examples.

**Quick Reference**:
- 0: Lavender, 1: Gold, 2: Cyan, 3: Burgundy
- 4: Green, 5: Brown, 6: Royal Blue, 7: Pumpkin
- 8: Gray, 9: Volt, 10: Teal, 11: Magenta
- 12: Mint, 13: Dark Brown, 14: Blue, 15: Pink

## Complete Configuration Example

```json
{
  "description": "My SSMS EnvTabs Configuration",
  "version": "1.2.0",
  "settings": {
    "enableAutoRename": true,
    "enableAutoColor": true,
    "autoConfigure": "server db",
    "enableConfigurePrompt": true
  },
  "groups": [
    {
      "groupName": "Prod-Main",
      "server": "PROD-SQL-01",
      "database": "CustomerDB",
      "priority": 5,
      "colorIndex": 3,
      "comment": "Most critical production database"
    },
    {
      "groupName": "Prod",
      "server": "%PROD%",
      "database": "%",
      "priority": 10,
      "colorIndex": 3,
      "comment": "All other production servers"
    },
    {
      "groupName": "QA",
      "server": "%QA%",
      "database": "%",
      "priority": 20,
      "colorIndex": 1
    },
    {
      "groupName": "Dev",
      "server": "%DEV%",
      "database": "%",
      "priority": 30,
      "colorIndex": 4
    },
    {
      "groupName": "Local",
      "server": "localhost",
      "database": "%",
      "priority": 40,
      "colorIndex": 2
    },
    {
      "groupName": "Test",
      "server": "%",
      "database": "%Test%",
      "priority": 50,
      "colorIndex": 9
    }
  ]
}
```

## Editing Tips

### JSON Syntax

- Strings must be in double quotes: `"value"`
- Numbers don't need quotes: `10`
- Booleans: `true` or `false` (no quotes)
- Arrays use square brackets: `[...]`
- Objects use curly braces: `{...}`
- Commas separate items, but no comma after the last item

### Validation

If you save invalid JSON:
- SSMS EnvTabs will log errors to `%LocalAppData%\SSMS EnvTabs\runtime.log`
- Tabs may not be renamed/colored until fixed
- Use a JSON validator or SSMS's JSON validation

### Testing Changes

1. Make a small change
2. Save the file
3. Open a new query tab or switch connections
4. Verify the tab is renamed/colored as expected
5. Adjust and repeat

## Advanced Scenarios

### Multiple Environments per Color

Use the same color for related environments:

```json
{
  "groupName": "Prod-US",
  "server": "%US-PROD%",
  "database": "%",
  "priority": 10,
  "colorIndex": 3
},
{
  "groupName": "Prod-EU",
  "server": "%EU-PROD%",
  "database": "%",
  "priority": 11,
  "colorIndex": 3
}
```

### Different Colors for Different Databases on Same Server

```json
{
  "groupName": "Finance",
  "server": "PROD-SQL-01",
  "database": "FinanceDB",
  "priority": 5,
  "colorIndex": 11
},
{
  "groupName": "HR",
  "server": "PROD-SQL-01",
  "database": "HRDB",
  "priority": 6,
  "colorIndex": 14
}
```

### Excluding Specific Databases

Use priority to exclude before a broader rule matches:

```json
{
  "groupName": "Test",
  "server": "DEV-SQL-01",
  "database": "TestDB",
  "priority": 5,
  "colorIndex": 9
},
{
  "groupName": "Dev",
  "server": "DEV-SQL-01",
  "database": "%",
  "priority": 10,
  "colorIndex": 4
}
```

## Next Steps

- [Wildcard Patterns](Wildcard-Patterns.md) - Learn more about pattern matching
- [Color Reference](Color-Reference.md) - See all available colors
- [Tips & Best Practices](Tips-and-Best-Practices.md) - Optimize your configuration
