# SSMS EnvTabs

A Visual Studio Extension for SQL Server Management Studio (SSMS) that automatically organizes and colors your query tabs based on server and database connections. Keep your production, QA, and development environments visually distinct and easily identifiable.

![SSMS Version](https://img.shields.io/badge/SSMS-22.0%2B-blue)
![.NET Framework](https://img.shields.io/badge/.NET-4.7.2%2B-purple)
![License](https://img.shields.io/badge/license-MIT-green)

## Key Features

- **Smart Tab Renaming** - Automatically prefix query tabs with environment names (e.g., "Prod-1", "QA-2")
- **Color-Coded Tabs** - 16 distinct colors to visually separate different environments
- **Flexible Rules** - Pattern matching with wildcards for server and database names
- **Auto-Configuration** - Automatically create rules for new connections
- **Priority-Based** - Control which rules match first with priority ordering

## Quick Start

### Installation

1. Download the latest `.vsix` file from [Releases](https://github.com/Blake-goofy/SSMS-EnvTabs/releases)
2. Close SSMS completely
3. Double-click the `.vsix` file and click **Install**
4. Launch SSMS and start working!

**Requirements:** SSMS 22.0+ and .NET Framework 4.7.2+

### Basic Usage

1. **Open Configuration**: In SSMS, go to **Tools** → **SSMS EnvTabs** → **Open Configuration**
2. **Edit Rules**: Configure your server/database patterns and colors
3. **Connect & Work**: Query tabs will automatically be renamed and colored based on your rules

Example configuration:
```json
{
  "settings": {
    "enableAutoRename": true,
    "enableAutoColor": true
  },
  "groups": [
    {
      "groupName": "Prod",
      "server": "%PROD%",
      "database": "%",
      "priority": 10,
      "colorIndex": 3
    }
  ]
}
```

## Documentation

For detailed documentation, please visit the [Wiki](wiki/):

- **[Installation Guide](wiki/Installation-Guide.md)** - Detailed installation instructions and troubleshooting
- **[Configuration Guide](wiki/Configuration-Guide.md)** - Complete configuration reference with examples
- **[Color Reference](wiki/Color-Reference.md)** - All 16 available colors
- **[Wildcard Patterns](wiki/Wildcard-Patterns.md)** - Pattern matching guide
- **[Hot it works](wiki/Hot-it-works.md)** - How EnvTabs uses SSMS color tabs and salt-based hashing
- **[Troubleshooting](wiki/Troubleshooting.md)** - Common issues and solutions
- **[Tips & Best Practices](wiki/Tips-and-Best-Practices.md)** - Optimize your workflow

## Contributing

Contributions are welcome! Please feel free to submit issues or pull requests.

### Building from Source

```bash
git clone https://github.com/Blake-goofy/SSMS-EnvTabs.git
cd SSMS-EnvTabs
# Open SSMS EnvTabs.sln in Visual Studio 2019+
# Build solution (requires Visual Studio SDK)
```

## Support

- **Issues**: [GitHub Issues](https://github.com/Blake-goofy/SSMS-EnvTabs/issues)
- **Discussions**: [GitHub Discussions](https://github.com/Blake-goofy/SSMS-EnvTabs/discussions)

## Author

**Blake Becker**

---

If you find this extension helpful, please consider giving it a star!
