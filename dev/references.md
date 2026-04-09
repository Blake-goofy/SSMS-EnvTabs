## Install / Uninstall (SSMS EnvTabs)

Open PowerShell, then:

Install:
```powershell
& "C:\Program Files\Microsoft SQL Server Management Studio 22\Release\Common7\IDE\VSIXInstaller.exe" "C:\Users\blake\source\repos\SSMS EnvTabs\SSMS EnvTabs\bin\release\SSMS EnvTabs.vsix"
```

Reinstall:
```
.\dev\reinstall.ps1
```

Uninstall:
```powershell
& "C:\Program Files\Microsoft SQL Server Management Studio 22\Release\Common7\IDE\VSIXInstaller.exe" /uninstall:SSMS_EnvTabs
```

---

## File Paths

**User Config:**
```
%USERPROFILE%\Documents\SSMS EnvTabs\TabGroupConfig.json
```

**Logs:**
- File log: `%LocalAppData%\SSMS EnvTabs\runtime.log`
- ActivityLog: `%AppData%\Microsoft\SQL Server Management Studio\22.0\ActivityLog.xml`

**ColorByRegex Config (SSMS temp):**
```
C:\Users\blake\AppData\Local\Temp\<guid>\ColorByRegexConfig.txt
```
- SSMS creates this after opening first query tab
- Extension writes generated regex lines here

## VS Code Build And Test

This repo includes VS Code tasks and helper scripts so you can build and run tests without leaving VS Code.

Requirements:

- Visual Studio or Build Tools with MSBuild
- .NET SDK (for running test project)

Run tasks from VS Code:

- `Terminal -> Run Task -> Build VSIX (Debug)`
- `Terminal -> Run Task -> Build VSIX (Release)`
- `Terminal -> Run Task -> Run Regression Tests`

Helper scripts:

- `dev/vscode-build.ps1`
- `dev/vscode-test.ps1`

Task definitions:

- `.vscode/tasks.json`