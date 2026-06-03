## Install / Uninstall (SSMS EnvTabs)

Open PowerShell, then:

Install:
```powershell
& "C:\Program Files\Microsoft SQL Server Management Studio 22\Release\Common7\IDE\VSIXInstaller.exe" "C:\Users\blake\source\repos\SSMS EnvTabs\SSMS EnvTabs\bin\release\SSMS EnvTabs.vsix"
```

Reinstall:
```
.\.vscode\build-install.ps1
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

## VS Code Build And Install

This repo includes a single VS Code task for the normal development loop: build the Release VSIX and install it into SSMS.

Requirements:

- Visual Studio or Build Tools with MSBuild
- SSMS 22 with `VSIXInstaller.exe`

Run tasks from VS Code:

- `Terminal -> Run Task -> Build and Install`

Helper scripts:

- `.vscode/build-install.ps1`

Task definitions:

- `.vscode/tasks.json`
