The configuration file is located at `%USERPROFILE%\Documents\SSMS EnvTabs\TabGroupConfig.json`.

Open it via **Tools** > **EnvTabs: Configure Rules...**.

## Settings

| Setting | Type | Description |
|---|---|---|
| `enableAutoRename` | Bool | Renames `SQLQueryX.sql` tabs to `GroupName-#` (e.g., `Prod-1`). |
| `enableAutoColor` | Bool | Colors tabs based on rules (requires SSMS 22.0+). |
| `autoConfigure` | String | Auto-create rules for new connections: `"off"`, `"server"`, `"server db"`. |
| `enableConfigurePrompt` | Bool | Prompt to edit config after auto-creating a rule. |
| `enableLogging` | Bool | Enables debug logging to helps troubleshooting issues. |

## Troubleshooting & Logging

If you encounter issues, you can inspect the runtime log.

- **Log File Location**: `%LocalAppData%\SSMS EnvTabs\runtime.log`
- **Enable Logging**: Ensure `"enableLogging": true` is set in your config `settings`.

### Auto-Configuration Prompt

When `enableConfigurePrompt` is true and a new rule is automatically created, you will see this prompt:

![Auto-Configuration Prompt](./images/config-prompt.png)

The rule is created immediately with default values before the prompt appears. The buttons allow you to modify or accept it:

- **Save**: Updates the rule with any changes you made to the Name or Color.
- **Cancel**: Closes the prompt. The rule remains active with its default values (Auto-generated Name and next available Color).
- **Open Config**: Saves your changes and immediately opens the `TabGroupConfig.json` file for advanced editing.

## Group Rules

Rules determine how tabs are grouped and colored.

| Field | Description |
|---|---|
| `groupName` | The prefix for the tab name (e.g., "Prod"). |
| `server` | Server name pattern (SQL LIKE syntax: `%` wildcard). |
| `database` | Database name pattern. |
| `priority` | Evaluation order (lowest matches first). |
| `colorIndex` | 0-15 (See [Color Reference](Color-Reference.md)). |

## Example

```json
{
  "groups": [
    {
      "groupName": "Prod",
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
    }
  ]
}
```
