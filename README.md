# SSMS EnvTabs

A Visual Studio Extension for SQL Server Management Studio (SSMS) that automatically names and colors your query tabs based on server and database connections. Keep your production, QA, and development environments visually distinct and easily identifiable.

![SSMS EnvTabs Demo](images/tab-colors.png)

![SSMS Version](https://img.shields.io/badge/SSMS-22.0%2B-blue)
![.NET Framework](https://img.shields.io/badge/.NET-4.7.2%2B-purple)
![License](https://img.shields.io/badge/license-MIT-green)

## Key Features

- **Smart Tab Renaming** - Automatically prefix query tabs with environment names (e.g., "Prod-1", "QA-2")
- **Color-Coded Tabs** - 16 distinct colors to visually separate different environments
- **Flexible Rules** - Pattern matching with wildcards for server and database names
- **Auto-Configuration** - Automatically create rules for new connections
- **Priority-Based** - Control which rules match first with priority ordering

## Getting Started

1.  **Install**: Download the latest `.vsix` from [GitHub Releases](https://github.com/Blake-goofy/SSMS-EnvTabs/releases) and run the installer.
2.  **Enable Coloring**: In SSMS, ensure *Tools > Options > Environment > Tabs and Windows > "Color document tabs by regular expression"* is selected.
3.  **Configure**: Go to **Tools > EnvTabs: Configure Rules...** to define your connection groups.

### Example Configuration

```json
{
  "settings": {
    "enableAutoRename": true,
    "enableAutoColor": true,
    "autoConfigure": "server db"
  },
  "groups": [
    {
      "groupName": "Prod",
      "server": "%PROD%",
      "priority": 10,
      "colorIndex": 3
    }
  ]
}
```

## Documentation

Full documentation is available in the [GitHub Wiki](https://github.com/Blake-goofy/SSMS-EnvTabs/wiki).

*   **[Installation Guide](https://github.com/Blake-goofy/SSMS-EnvTabs/wiki/Installation-Guide)** - Setup instructions and troubleshooting.
*   **[Configuration Guide](https://github.com/Blake-goofy/SSMS-EnvTabs/wiki/Configuration-Guide)** - Complete reference for settings and rules.
*   **[Color Reference](https://github.com/Blake-goofy/SSMS-EnvTabs/wiki/Color-Reference)** - Visual guide for the 16 available colors.
*   **[Wildcard Patterns](https://github.com/Blake-goofy/SSMS-EnvTabs/wiki/Wildcard-Patterns)** - How to use SQL-like wildcards for matching.
*   **[How it works](https://github.com/Blake-goofy/SSMS-EnvTabs/wiki/How-it-works)** - Technical deep dive into the coloring mechanism.
*   **[Troubleshooting](https://github.com/Blake-goofy/SSMS-EnvTabs/wiki/Troubleshooting)** - Common issues and solutions.

## Support & Contributing

*   **Issues**: [GitHub Issues](https://github.com/Blake-goofy/SSMS-EnvTabs/issues)
*   **Discussions**: [GitHub Discussions](https://github.com/Blake-goofy/SSMS-EnvTabs/discussions)
*   **Source Code**: To build from source, clone the repo and open `SSMS EnvTabs.sln` in Visual Studio 2019+ (VS SDK required).

## Author

**Blake Becker**

---

If you find this extension helpful, please consider giving it a star!
