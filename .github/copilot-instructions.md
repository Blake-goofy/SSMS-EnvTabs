# SSMS EnvTabs - AI Coding Instructions

## Overview
SSMS EnvTabs is a Visual Studio Extension (VSIX) for SQL Server Management Studio (SSMS) 2017+ (shell-based). It groups and colors query tabs based on connection properties (Server/Database) defined in user configuration.

## Architecture & Key Components
- **Entry Point**: `SSMS_EnvTabsPackage.cs` initializes the extension as an `AsyncPackage`.
- **Event Loop**: `RdtEventManager.cs` subscribes to `IVsRunningDocTableEvents` to detect when query windows are opened, closed, or switched. This is the core "engine" of the extension.
- **Logic**:
  - `TabRuleMatcher.cs`: Matches current connection info against loaded rules.
  - `TabRenamer.cs`: Renames window captions (e.g., "Prod-1") using `IVsWindowFrame.SetProperty`.
  - `ColorByRegexConfigWriter.cs`: Writes regex-based coloring rules to a temp file (`ColorByRegexConfig.txt`) consumed by SSMS/MIDS for coloring query tabs.
- **Configuration**: User rules are loaded from `%USERPROFILE%\Documents\SSMS EnvTabs\TabGroupConfig.json` via `TabGroupConfigLoader.cs`.

### Coloring Logic & Challenges
- **Mechanism**: The extension writes a `ColorByRegexConfig.txt` file which SSMS reads.
- **Regex Format**: We generate regexes based on **Filenames only** (e.g., `(query.sql|script.sql)`) to avoid path dependency issues when files move.
- **Color Assignment**:
  - SSMS assigns colors based on a **hash of the regex string itself**.
  - Changing the regex (even adding whitespace) changes the assigned color unpredictably.
  - The extension supports a `salt` property in `TabGroupConfig.json` (e.g., `"salt": "prod-1"`) which is appended as a regex comment `(?#salt:prod-1)`. This allows the user to manually "roll" for a different color by changing the salt string without affecting the matching logic.
  - Line number or order in the config file does **not** determine the color.
- **Color Persistence**:
  - SSMS outputs a log-like file `customized-groupid-color-<GUID>.json` mapping internal GroupIDs to ColorIndexes. This file is for reference/output only; writing to it has no effect.
  - **Challenge**: There is currently no known programmatic way to force a specific color (e.g., "Red") for a given rule. The "Set Tab Color" UI feature in SSMS is the only known manual override.
  - **Future Goal**: Investigate if the "Set Tab Color" command can be invoked programmatically or if the regex string can be "salted" to target specific hashes/colors.

## Developer Workflows

### Debugging
- **Start Action**: The project defaults to `devenv.exe /rootsuffix Exp`. To debug in SSMS, change project properties → Debug → Start external program to your `ssms.exe` path (e.g., `C:\Program Files (x86)\Microsoft SQL Server Management Studio 18\Common7\IDE\Ssms.exe`).
- **Command Arguments**: Use `/log` to enable SSMS logging if needed.
- **Attaching**: You can attach to a running `Ssms.exe` process if the extension is installed.

### Logging
- **Runtime Log**: The extension writes its own log to `%LocalAppData%\SSMS EnvTabs\runtime.log` via `EnvTabsLog.cs`. Check this first for logic errors. **Note: This file logging is for development only and will be removed in the final release.**
- **ActivityLog**: Standard VS activity log is at `%AppData%\Microsoft\SQL Server Management Studio\...\ActivityLog.xml`.

### Build & Deploy
- **Build**: Uses standard MSBuild with VS SDK targets.
- **Install**: Use `VSIXInstaller.exe` manually or via script (see `dev\references.md`) to install the `.vsix` into SSMS.

## Project Conventions

### Threading
- **Strict UI Thread Usage**: Most VS Shell interactions (RDT, Window Frames) must happen on the UI thread.
- **Pattern**: Use `await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync()` or `ThreadHelper.ThrowIfNotOnUIThread()` before accessing `IVs*` services.

### Configuration Pattern
- **External Config**: Unlike typical VS extensions using `DialogPage`, this project uses a standalone JSON file in the User's Documents folder to allow sharing/scripting of rules.
- **Reloading**: Config is reloaded automatically or on specific triggers (check `RdtEventManager.cs` for logic).

### VS Interop
- **Services**: `SVsRunningDocumentTable`, `SVsUIShellOpenDocument` are key services.
- **Properties**: Captions are modified via `VSFPROPID_OwnerCaption` or fallback to `VSFPROPID_Caption`.
