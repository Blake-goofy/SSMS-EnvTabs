# Installation Guide

## Prerequisites

- **SSMS v22.0+**: [Download here](https://docs.microsoft.com/sql/ssms/download-sql-server-management-studio-ssms)
- **.NET Framework 4.7.2+**: Included with Windows 10/11.

## Installation

### 1. VSIX Installer (Recommended)

1. Download the latest `.vsix` from [Releases](https://github.com/Blake-goofy/SSMS-EnvTabs/releases).
2. **Close SSMS**.
3. Run the `.vsix` installer.
4. Click **Install**.
5. Restart SSMS.

### 2. Manual Installation

1. Open `%LocalAppData%\Microsoft\SQL Server Management Studio\<version>\Extensions`.
2. Extract the VSIX (it's a zip file) into a new subfolder.
3. Restart SSMS.

## Post-Installation Configuration

**IMPORTANT:** For tab coloring to work, you must enable "Color document tabs by regular expression" in SSMS options.

1. Go to **Tools** > **Options**.
2. Navigate to **Environment** > **Tabs and Windows**.
3. Check **Color document tabs by regular expression**.

OR use the settings within the document window:

![Enable Regex Coloring](./images/regular-expressions-tabs.png)

## Verification

1. Go to **Tools** > **EnvTabs: Configure Rules...**.
2. If the configuration file opens, the extension is active.
