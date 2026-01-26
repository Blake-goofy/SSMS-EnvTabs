## Tabs Not Renaming

1.  **Check `enableAutoRename`**: Ensure it is `true` in config.
2.  **Unsaved Files Only**: Only `SQLQueryX.sql` tabs are renamed to prevent overwriting filenames.
3.  **Check Rules**: Ensure your server/database matches a rule. Try a catch-all (`%`) rule with priority 100.
4.  **Priorities**: Lower numbers match first.
5.  **JSON Syntax**: Validate your JSON file for errors.

## Colors Not Showing

1.  **Check `enableAutoColor`**: Ensure it is `true`.
2.  **Check `Color tabs by`** setting in SSMS: Ensure it is `Regular expression`.
3.  **SSMS Version**: Requires SSMS 22.0+.
4.  **Restart**: Restart SSMS after enabling coloring.
5.  **Valid Index**: Ensure `colorIndex` is 0-15.

## Issues?

Check the log at `%LocalAppData%\SSMS EnvTabs\runtime.log`.
