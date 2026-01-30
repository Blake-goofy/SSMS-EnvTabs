# SSMS EnvTabs

A Visual Studio Extension for SQL Server Management Studio (SSMS) that automatically names query tabs and leverages SSMS's native regex feature to color them based on server and database connections. Keep your production, QA, and development environments visually distinct and easily identifiable.

![SSMS EnvTabs Demo](images/tab-colors.png)

![SSMS Version](https://img.shields.io/badge/SSMS-22.0%2B-blue)
![.NET Framework](https://img.shields.io/badge/.NET-4.7.2%2B-purple)
![License](https://img.shields.io/badge/license-MIT-green)

## Key Features

- **Color-Coded Tabs** - 16 distinct colors to visually separate different environments
- **Tab Renaming** - Automatically name query tabs with environment names (e.g., "1. Prod", "1. QA")
- **Auto-Configuration** - Automatically create rules for new connections (retained between sessions)
- **Manual-Regex** - Add your own regex that you want applied to the ColorByRegexConfig.txt file (retained between sessions)

## Getting Started

1.  **Install**: Download the latest `.vsix` from [GitHub Releases](https://github.com/Blake-goofy/SSMS-EnvTabs/releases) and run the installer.
2.  **Enable Coloring**: In SSMS, ensure *Tools > Options > Environment > Tabs and Windows > "Color document tabs by regular expression"* is selected.
3.  **Enjoy**: That's it, just open a new query and you will be guided to assign a name and color to the new connection.

### On-Demand Configuration

When you connect to a server or database that doesn't have a matching rule, EnvTabs will prompt you to configure it. This ensures you only create rules for the connections you actually use!

![Configuration Prompt](images/config-prompt-colors.png)

### Example Configuration

```json
{
  "settings": {
    "enableLogging": false,
    "enableAutoRename": true,
    "enableAutoColor": true,
    "enableConfigurePrompt": true,
    "enableConnectionPolling": true,
    "autoConfigure": "server db",
    "newQueryRenameStyle": "[#]. [groupName]"
  },
  "serverAlias": {
    "MY-APP-SERVER": "AppServer"
  },
  "connectionGroups": [
    {
      "groupName": "Example: Exact Match",
      "server": "MY-APP-SERVER",
      "database": "MyDatabase",
      "priority": 10,
      "colorIndex": 9
    }
  ]
}

```
## Documentation

Full documentation is available in the [GitHub Wiki](https://github.com/Blake-goofy/SSMS-EnvTabs/wiki).

- **[Installation](https://github.com/Blake-goofy/SSMS-EnvTabs/wiki/Installation-Guide)**: Setup and requirements.
- **[Configuration](https://github.com/Blake-goofy/SSMS-EnvTabs/wiki/Configuration-Guide)**: Global settings and prompts.
- **[Group Rules](https://github.com/Blake-goofy/SSMS-EnvTabs/wiki/Group-Rules)**: Creating rules for your servers.
- **[Manual Regex Rules](https://github.com/Blake-goofy/SSMS-EnvTabs/wiki/Manual-Regex)**: Advanced regex matching overrides.
- **[Tab Text Settings](https://github.com/Blake-goofy/SSMS-EnvTabs/wiki/Tab-Text-Settings)**: How to clean up tab text.
- **[Color Reference](https://github.com/Blake-goofy/SSMS-EnvTabs/wiki/Color-Reference)**: List of available colors (0-15).
- **[Wildcard Patterns](https://github.com/Blake-goofy/SSMS-EnvTabs/wiki/Wildcard-Patterns)**: Matching syntax.
- **[How it Works](https://github.com/Blake-goofy/SSMS-EnvTabs/wiki/How-it-works)**: Technical details.
- **[Troubleshooting](https://github.com/Blake-goofy/SSMS-EnvTabs/wiki/Troubleshooting)**: Common issues and solutions.

## Support & Contributing

*   **Issues**: [GitHub Issues](https://github.com/Blake-goofy/SSMS-EnvTabs/issues)
*   **Discussions**: [GitHub Discussions](https://github.com/Blake-goofy/SSMS-EnvTabs/discussions)
*   **Source Code**: To build from source, clone the repo and open `SSMS EnvTabs.sln` in Visual Studio 2019+ (VS SDK required).

## Author

**Blake Becker**

---

If you find this extension helpful, please consider giving it a star!
