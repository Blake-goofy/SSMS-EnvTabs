# SSMS EnvTabs

A Visual Studio Extension for SQL Server Management Studio (SSMS) that automatically groups and colors query tabs based on their server and database connections.

## Features

- **Automatic Tab Renaming**: Query tabs are renamed with custom prefixes based on their connection (e.g., "Prod-1", "QA-2").
- **Color-Coded Tabs**: Tabs are automatically colored based on connection properties, making it easy to distinguish between environments at a glance.
- **Flexible Configuration**: Define custom rules using server/database patterns with wildcard support.
- **Priority-Based Matching**: Rules are evaluated in priority order for precise control.
- **Auto-Configuration**: Optionally create rules automatically when connecting to new servers/databases.
- **16 Color Options**: Choose from 16 distinct colors to organize your workspace.

## Installation

### Prerequisites

- **SQL Server Management Studio (SSMS)** version 22.0 or later
- **Microsoft .NET Framework** 4.7.2 or higher

### Installing the Extension

1. **Download the Extension**
   - Download the latest `.vsix` file from the [Releases page](https://github.com/Blake-goofy/SSMS-EnvTabs/releases)

2. **Close SSMS**
   - Make sure SQL Server Management Studio is completely closed before installation

3. **Install the Extension**
   - Double-click the downloaded `.vsix` file
   - The Visual Studio Extension Installer will open
   - Click **Install** and wait for the installation to complete
   - Click **Close** when finished

4. **Launch SSMS**
   - Open SQL Server Management Studio
   - The extension will be active and ready to use

### Verifying Installation

1. Open SSMS
2. Go to **Tools** → **Extensions and Updates** (or **Manage Extensions** in newer versions)
3. Look for **SSMS_EnvTabs** in the installed extensions list

## How to Use

### First-Time Setup

When you first install SSMS EnvTabs, a default configuration file is created at:

```
%USERPROFILE%\Documents\SSMS EnvTabs\TabGroupConfig.json
```

### Opening the Configuration File

There are two ways to open and edit your configuration:

1. **From SSMS Menu** (Recommended)
   - Open SSMS
   - Go to **Tools** → **SSMS EnvTabs** → **Open Configuration**
   - The configuration file will open in SSMS

2. **Manually**
   - Navigate to `%USERPROFILE%\Documents\SSMS EnvTabs\` in File Explorer
   - Open `TabGroupConfig.json` in any text editor

### Basic Configuration

The configuration file uses JSON format. Here's what a basic setup looks like:

```json
{
  "settings": {
    "enableAutoRename": true,
    "enableAutoColor": true,
    "autoConfigure": "server db",
    "enableConfigurePrompt": true
  },
  "groups": [
    {
      "groupName": "Production",
      "server": "%PROD%",
      "database": "%",
      "priority": 10,
      "colorIndex": 3
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
    }
  ]
}
```

### Configuration Options

#### Settings

- **enableAutoRename** (`true`/`false`): Enables automatic tab renaming. Only affects unsaved queries (SQLQueryX.sql). Saved files keep their original names.
- **enableAutoColor** (`true`/`false`): Enables automatic tab coloring based on matched rules.
- **autoConfigure** (`"off"`, `"server"`, `"server db"`): 
  - `"off"`: Disabled - no automatic rule creation
  - `"server"`: Create rules grouped by server name only
  - `"server db"`: Create rules grouped by both server and database
- **enableConfigurePrompt** (`true`/`false`): Shows a prompt to edit the config when new rules are automatically added.

#### Group Rules

Each group rule defines how tabs should be renamed and colored:

- **groupName**: The prefix shown on matching tabs (e.g., "Prod-1", "QA-2")
- **server**: Server name pattern (supports SQL LIKE wildcards with `%`)
- **database**: Database name pattern (supports SQL LIKE wildcards with `%`)
- **priority**: Lower numbers are checked first (e.g., 10 before 20)
- **colorIndex**: A number from 0-15 representing the tab color (see color reference below)

#### Available Colors (colorIndex values)

| Index | Color Name  | Index | Color Name    |
|-------|-------------|-------|---------------|
| 0     | Lavender    | 8     | Gray          |
| 1     | Gold        | 9     | Volt          |
| 2     | Cyan        | 10    | Teal          |
| 3     | Burgundy    | 11    | Magenta       |
| 4     | Green       | 12    | Mint          |
| 5     | Brown       | 13    | Dark Brown    |
| 6     | Royal Blue  | 14    | Blue          |
| 7     | Pumpkin     | 15    | Pink          |

### Using Wildcards

The `server` and `database` fields support SQL LIKE pattern matching:

- `%` matches any sequence of characters
- `_` matches any single character

**Examples:**

```json
{
  "groupName": "All Production",
  "server": "%PROD%",           // Matches: SERVER-PROD-01, PROD-DB, MY-PROD-SERVER
  "database": "%",              // Matches any database
  "priority": 10,
  "colorIndex": 3
}
```

```json
{
  "groupName": "Specific DB",
  "server": "MY-SERVER",        // Exact match: MY-SERVER only
  "database": "CustomerDB",     // Exact match: CustomerDB only
  "priority": 5,
  "colorIndex": 14
}
```

```json
{
  "groupName": "Test Databases",
  "server": "%",                // Any server
  "database": "Test%",          // Matches: TestDB, Test_Customer, Testing
  "priority": 50,
  "colorIndex": 1
}
```

### How Tab Matching Works

1. When you open a query window or switch connections, SSMS EnvTabs reads your current server and database
2. Rules are evaluated in **priority order** (lowest number first)
3. The **first matching rule** wins
4. If `enableAutoRename` is enabled, the tab is renamed with the `groupName` + a counter (e.g., "Prod-1", "Prod-2")
5. If `enableAutoColor` is enabled, the tab is colored according to the `colorIndex`

### Example Workflow

1. **Connect to a server** in SSMS (e.g., "PROD-SQL-01", database "CustomerDB")
2. **Open a new query** (creates "SQLQuery1.sql" tab)
3. **Extension activates**:
   - Finds matching rule (e.g., server pattern "%PROD%")
   - Renames tab to "Prod-1"
   - Colors tab with Burgundy (colorIndex 3)
4. **Open another query** on the same connection
   - Renamed to "Prod-2" (same group, incremented counter)
   - Same Burgundy color

## Auto-Configuration Feature

When `autoConfigure` is set to `"server"` or `"server db"`, the extension will:

1. Detect when you connect to a server/database that doesn't match any existing rule
2. Automatically create a new rule for that connection
3. (Optional) Prompt you to review and edit the configuration

This is useful for quickly building up your rule set as you work with different environments.

## Troubleshooting

### Tabs Aren't Being Renamed

- **Check `enableAutoRename` setting**: Make sure it's set to `true` in your config
- **Saved files aren't renamed**: Only unsaved queries (SQLQueryX.sql) are renamed. This is by design to preserve your file names
- **Check your rules**: Ensure you have rules that match your server/database patterns
- **Check priority order**: A higher-priority rule might be matching first

### Colors Aren't Showing

- **Check `enableAutoColor` setting**: Make sure it's set to `true` in your config
- **Restart SSMS**: Sometimes color changes require SSMS to be restarted
- **Check colorIndex values**: Must be between 0-15

### Configuration File Not Found

If the configuration file is missing:

1. Open SSMS
2. Go to **Tools** → **SSMS EnvTabs** → **Open Configuration**
3. The extension will automatically create the default configuration file

### Extension Not Loading

- **Check SSMS Version**: Requires SSMS 22.0 or later
- **Check .NET Framework**: Requires .NET Framework 4.7.2 or higher
- **Reinstall the Extension**: Uninstall via **Tools** → **Extensions and Updates**, then reinstall the `.vsix` file

### Viewing Logs

For troubleshooting, logs are written to:

```
%LocalAppData%\SSMS EnvTabs\runtime.log
```

You can check this file if you encounter issues with tab renaming or coloring.

## Tips and Best Practices

1. **Use Lower Priority Numbers for Specific Rules**: Put your most specific rules (exact server/database matches) at lower priorities (e.g., 10) and general wildcard rules at higher priorities (e.g., 50+)

2. **Group Related Environments**: Use similar colors for related environments (e.g., all production environments in red tones)

3. **Test Your Patterns**: Use the `%` wildcard carefully - overly broad patterns might match unintended servers

4. **Keep GroupNames Short**: Short group names (4-8 characters) work best in the tab UI

5. **Document Your Colors**: Add meaningful groupName values so you remember what each color represents

## Uninstalling

To remove SSMS EnvTabs:

1. Close SSMS
2. Open **Control Panel** → **Programs and Features**
3. Find **SSMS_EnvTabs** in the list
4. Click **Uninstall**
5. (Optional) Manually delete the configuration folder at `%USERPROFILE%\Documents\SSMS EnvTabs\`

## Building from Source

If you want to build the extension yourself:

1. Clone the repository
2. Open `SSMS EnvTabs.sln` in Visual Studio 2019 or later
3. Build the solution (requires Visual Studio SDK)
4. The `.vsix` file will be generated in the `bin\Debug` or `bin\Release` folder

## License

[Add your license information here]

## Support

For issues, feature requests, or questions:
- Open an issue on [GitHub Issues](https://github.com/Blake-goofy/SSMS-EnvTabs/issues)
- Check existing issues for solutions to common problems

## Credits

Created by Blake Becker
