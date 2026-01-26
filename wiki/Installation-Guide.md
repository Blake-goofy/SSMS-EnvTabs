# Installation Guide

This guide provides detailed instructions for installing SSMS EnvTabs.

## Prerequisites

Before installing SSMS EnvTabs, ensure you have:

- **SQL Server Management Studio (SSMS)** version 22.0 or later
  - Released November 2025
  - Download from [Microsoft's official site](https://docs.microsoft.com/sql/ssms/download-sql-server-management-studio-ssms)
- **Microsoft .NET Framework** 4.7.2 or higher
  - Usually included with Windows 10/11
  - Download from [Microsoft](https://dotnet.microsoft.com/download/dotnet-framework)

### Checking Your SSMS Version

1. Open SSMS
2. Go to **Help** → **About Microsoft SQL Server Management Studio**
3. Check the version number in the dialog

## Installation Steps

### Method 1: Using the VSIX Installer (Recommended)

1. **Download the Extension**
   - Visit the [Releases page](https://github.com/Blake-goofy/SSMS-EnvTabs/releases)
   - Download the latest `.vsix` file (e.g., `SSMS_EnvTabs-1.0.vsix`)
   - Save it to a location you can easily find (e.g., Downloads folder)

2. **Close SSMS Completely**
   - Make sure all SSMS windows are closed
   - Check Task Manager to ensure no `Ssms.exe` processes are running
   - This is important - the installation will fail if SSMS is running

3. **Run the Installer**
   - Double-click the downloaded `.vsix` file
   - The Visual Studio Extension Installer will open
   - Review the installation details

4. **Install the Extension**
   - Click the **Install** button
   - Wait for the installation to complete (usually takes 10-30 seconds)
   - You'll see a green checkmark when installation is successful
   - Click **Close** to exit the installer

5. **Launch SSMS**
   - Open SQL Server Management Studio
   - The extension will load automatically
   - You may see a brief initialization message

### Method 2: Manual Installation (Advanced)

For advanced users who need to install manually:

1. Locate your SSMS extensions folder:
   ```
   %LocalAppData%\Microsoft\SQL Server Management Studio\<version>\Extensions
   ```

2. Extract the `.vsix` file (it's a ZIP archive) to a subfolder in the Extensions directory

3. Restart SSMS

## Verifying Installation

After installation, verify that SSMS EnvTabs is loaded:

### Check Extension Manager

1. Open SSMS
2. Go to **Tools** → **Extensions and Updates** (or **Manage Extensions** in newer versions)
3. Look for **SSMS_EnvTabs** in the **Installed** section
4. It should show as enabled with version information

### Check Tools Menu

1. In SSMS, go to **Tools** menu
2. You should see a new **SSMS EnvTabs** submenu
3. The submenu should contain **Open Configuration** option

### Test Basic Functionality

1. Go to **Tools** → **SSMS EnvTabs** → **Open Configuration**
2. A configuration file should open in SSMS
3. If you see the JSON configuration file, installation was successful!

## First-Time Configuration

After installation, the extension creates a default configuration file at:

```
%USERPROFILE%\Documents\SSMS EnvTabs\TabGroupConfig.json
```

This file is created automatically the first time you:
- Open the configuration from the Tools menu, or
- Open a query window (if auto-configuration is enabled)

See the [Quick Start Guide](Quick-Start.md) for next steps.

## Troubleshooting Installation

### "This extension is not installable" Error

**Cause**: SSMS version is too old

**Solution**: 
- Upgrade to SSMS 22.0 or later
- Check your SSMS version in **Help** → **About**

### Installation Hangs or Fails

**Possible causes and solutions**:

1. **SSMS is still running**
   - Open Task Manager (Ctrl+Shift+Esc)
   - Look for `Ssms.exe` under Processes
   - End any SSMS processes
   - Try installation again

2. **Insufficient permissions**
   - Right-click the `.vsix` file
   - Select **Run as administrator**
   - Try installation again

3. **Corrupted download**
   - Re-download the `.vsix` file
   - Verify the file size matches the release page
   - Try installation again

### Extension Not Showing in Tools Menu

1. **Verify installation**:
   - Go to **Tools** → **Extensions and Updates**
   - Check if SSMS_EnvTabs is listed
   
2. **If listed but disabled**:
   - Enable it in Extensions and Updates
   - Restart SSMS

3. **If not listed at all**:
   - Reinstall the extension
   - Check the Activity Log:
     ```
     %AppData%\Microsoft\SQL Server Management Studio\<version>\ActivityLog.xml
     ```
   - Look for errors related to SSMS_EnvTabs

### Configuration File Not Created

If the configuration file doesn't appear:

1. Try opening it manually from **Tools** → **SSMS EnvTabs** → **Open Configuration**
2. If that fails, manually create the directory:
   ```
   %USERPROFILE%\Documents\SSMS EnvTabs\
   ```
3. Restart SSMS and try again

## Updating the Extension

To update to a newer version:

1. **Uninstall the old version** (see below)
2. **Install the new version** (follow installation steps above)
3. Your configuration file will be preserved

## Uninstalling

To remove SSMS EnvTabs:

### Method 1: Using Extensions Manager

1. Open SSMS
2. Go to **Tools** → **Extensions and Updates**
3. Find **SSMS_EnvTabs** in the Installed section
4. Click **Uninstall**
5. Restart SSMS to complete uninstallation

### Method 2: Using Control Panel

1. Close SSMS completely
2. Open **Control Panel** → **Programs and Features**
3. Find **SSMS_EnvTabs** in the list
4. Click **Uninstall**
5. Follow the prompts

### Removing Configuration Files

The uninstaller does not remove your configuration files. To completely remove all traces:

1. Delete the configuration folder:
   ```
   %USERPROFILE%\Documents\SSMS EnvTabs\
   ```

2. (Optional) Delete the log folder:
   ```
   %LocalAppData%\SSMS EnvTabs\
   ```

## Next Steps

- [Quick Start Guide](Quick-Start.md) - Get started in 5 minutes
- [Configuration Guide](Configuration-Guide.md) - Learn about all configuration options
- [Color Reference](Color-Reference.md) - Choose your tab colors
