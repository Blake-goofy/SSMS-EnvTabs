# Troubleshooting

Common issues and their solutions.

## Tabs Not Being Renamed

### Issue: Tabs remain as "SQLQuery1.sql"

**Possible Causes:**

1. **Auto-rename is disabled**
   - **Solution**: Check your configuration file
   - Ensure `"enableAutoRename": true` in the settings section
   - Save the file and open a new query tab

2. **File is already saved**
   - **Expected behavior**: Only unsaved query tabs (SQLQueryX.sql) are renamed
   - Saved files keep their original names to preserve your work
   - **Solution**: This is by design, not a bug

3. **No matching rule**
   - **Solution**: Check your group rules
   - Ensure you have a rule that matches your server/database
   - Add a catch-all rule with priority 100 to test:
     ```json
     {
       "groupName": "Test",
       "server": "%",
       "database": "%",
       "priority": 100,
       "colorIndex": 9
     }
     ```

4. **Higher priority rule is matching first**
   - **Solution**: Check rule priorities
   - Lower numbers are evaluated first
   - A general rule at priority 10 will match before a specific rule at priority 20

5. **Configuration file has syntax errors**
   - **Solution**: Validate your JSON
   - Check for missing commas, quotes, or brackets
   - Look at the log file: `%LocalAppData%\SSMS EnvTabs\runtime.log`

### Issue: Some tabs renamed, others not

**Cause**: Different connections matching different rules (or no rules)

**Solution**:
1. Check which server/database each tab is connected to
2. Verify your rules cover all your connections
3. Add a catch-all rule at the end:
   ```json
   {
     "groupName": "Other",
     "server": "%",
     "database": "%",
     "priority": 999,
     "colorIndex": 8
   }
   ```

## Colors Not Showing

### Issue: Tabs are renamed but not colored

**Possible Causes:**

1. **Auto-color is disabled**
   - **Solution**: Check your configuration
   - Set `"enableAutoColor": true` in settings
   - Save and restart SSMS

2. **SSMS version doesn't support coloring**
   - **Solution**: Verify SSMS version
   - Requires SSMS 22.0 or later
   - Check **Help** → **About** in SSMS

3. **Need to restart SSMS**
   - **Solution**: 
   - Close all SSMS windows
   - Restart SSMS
   - Colors should appear on new query tabs

4. **ColorIndex out of range**
   - **Solution**: Check colorIndex values
   - Must be between 0 and 15 (inclusive)
   - Check the log file for errors

### Issue: Wrong colors appearing

**Cause**: Incorrect colorIndex or wrong rule matching

**Solution**:
1. Verify the colorIndex in your matched rule
2. Check [Color Reference](Color-Reference.md) for correct index values
3. Ensure the rule you expect is actually matching (check priority)

## Extension Not Loading

### Issue: Extension menu not appearing

**Check Installation:**

1. **Verify extension is installed**
   - Open SSMS
   - Go to **Tools** → **Extensions and Updates**
   - Look for **SSMS_EnvTabs** in the Installed section

2. **Check if extension is enabled**
   - In Extensions and Updates, ensure it's not disabled
   - If disabled, enable it and restart SSMS

3. **Verify SSMS version**
   - Extension requires SSMS 22.0 or later
   - Check **Help** → **About** in SSMS
   - If too old, upgrade SSMS

4. **Check .NET Framework version**
   - Requires .NET Framework 4.7.2 or higher
   - Check in Control Panel → Programs and Features
   - Look for "Microsoft .NET Framework" entries

### Issue: Extension was working, now stopped

**Possible Causes:**

1. **Extension was disabled or uninstalled**
   - **Solution**: Check Extensions and Updates
   - Reinstall if necessary

2. **Configuration file corrupted**
   - **Solution**: 
   - Back up your config file
   - Delete or rename `TabGroupConfig.json`
   - Restart SSMS (creates new default config)
   - Merge your settings back

3. **SSMS update changed behavior**
   - **Solution**:
   - Check for extension updates
   - Review release notes
   - Report issue on GitHub

## Configuration File Issues

### Issue: Configuration file not found

**Solution:**

1. Try opening from SSMS menu:
   - **Tools** → **SSMS EnvTabs** → **Open Configuration**
   - Extension will create it if missing

2. Manually create the directory:
   - Create: `%USERPROFILE%\Documents\SSMS EnvTabs\`
   - Restart SSMS

3. Check permissions:
   - Ensure you can write to Documents folder
   - Run SSMS as administrator if needed (not recommended long-term)

### Issue: Changes not taking effect

**Solutions:**

1. **Save the file** (Ctrl+S in SSMS)
2. **Close and reopen query tabs** to see changes
3. **Check for JSON syntax errors**:
   - Missing or extra commas
   - Missing quotes around strings
   - Mismatched brackets
4. **Check the log file** for parsing errors:
   ```
   %LocalAppData%\SSMS EnvTabs\runtime.log
   ```

### Issue: JSON syntax errors

**Common mistakes:**

1. **Missing comma between items**
   ```json
   // WRONG
   {
     "groupName": "Prod"
     "server": "%PROD%"
   }
   
   // CORRECT
   {
     "groupName": "Prod",
     "server": "%PROD%"
   }
   ```

2. **Extra comma after last item**
   ```json
   // WRONG
   {
     "groupName": "Prod",
     "server": "%PROD%",
   }
   
   // CORRECT
   {
     "groupName": "Prod",
     "server": "%PROD%"
   }
   ```

3. **Single quotes instead of double quotes**
   ```json
   // WRONG
   'groupName': 'Prod'
   
   // CORRECT
   "groupName": "Prod"
   ```

4. **Comments in JSON** (not standard JSON)
   ```json
   // This may not work in all parsers
   {
     "groupName": "Prod", // This is production
     "server": "%PROD%"
   }
   ```

**Solution**: Use a JSON validator or SSMS's built-in JSON validation

## Pattern Matching Issues

### Issue: Pattern not matching expected servers

**Debugging steps:**

1. **Test with a catch-all rule**:
   ```json
   {
     "groupName": "Test",
     "server": "%",
     "database": "%",
     "priority": 5,
     "colorIndex": 9
   }
   ```
   If this works, your other patterns are wrong

2. **Check server name case**:
   - Server matching is usually case-insensitive
   - But try matching the exact case of your server name

3. **Check for extra spaces or special characters**:
   - Server name might have spaces
   - Try: `"server": "% PROD%"` (note the space)

4. **Use more specific patterns**:
   ```json
   // Instead of
   "server": "PROD"
   
   // Try
   "server": "%PROD%"
   ```

5. **Check priority order**:
   - A higher-priority rule might be matching first
   - Lower priority numbers = checked first

### Issue: Wildcard % not working

**Remember**: SQL LIKE syntax, not regex

```json
// WRONG (regex syntax)
"server": "PROD.*"

// CORRECT (SQL LIKE syntax)
"server": "PROD%"
```

```json
// WRONG (wildcard in middle without %)
"server": "PROD-01"

// CORRECT (exact match)
"server": "PROD-SQL-01"

// CORRECT (wildcard)
"server": "PROD-%"
```

## Performance Issues

### Issue: SSMS slow when opening tabs

**Possible Causes:**

1. **Too many complex rules**
   - **Solution**: Simplify rules
   - Use fewer, more general patterns
   - Combine similar rules

2. **Network latency**
   - Extension queries server/database names
   - **Solution**: Not much can be done; this is SSMS behavior

3. **Excessive logging**
   - Development logs may slow things down
   - **Note**: File logging will be removed in production release

## Viewing Logs

For detailed troubleshooting, check the log file:

```
%LocalAppData%\SSMS EnvTabs\runtime.log
```

**How to access:**
1. Press `Win+R`
2. Type: `%LocalAppData%\SSMS EnvTabs\`
3. Press Enter
4. Open `runtime.log` in Notepad

**What to look for:**
- Parse errors (configuration file problems)
- Pattern matching details
- Extension loading issues

**Note**: Logging is for development and will be removed in future releases.

## Getting Help

If you can't resolve the issue:

1. **Check existing issues**: [GitHub Issues](https://github.com/Blake-goofy/SSMS-EnvTabs/issues)
2. **Search closed issues**: Your problem may already be solved
3. **Create a new issue** with:
   - SSMS version
   - Extension version
   - Your configuration file (remove sensitive server names)
   - Steps to reproduce
   - Relevant log entries

## Common Configuration Mistakes

### Mistake 1: Rules in wrong order

```json
// WRONG - general rule before specific
{
  "groups": [
    {
      "groupName": "All",
      "server": "%",
      "database": "%",
      "priority": 10,
      "colorIndex": 8
    },
    {
      "groupName": "Prod",
      "server": "%PROD%",
      "database": "%",
      "priority": 20,
      "colorIndex": 3
    }
  ]
}
```

**Problem**: "All" rule matches everything first; "Prod" never matches

**Solution**: Reverse priorities:
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
      "groupName": "All",
      "server": "%",
      "database": "%",
      "priority": 100,
      "colorIndex": 8
    }
  ]
}
```

### Mistake 2: Expecting saved files to be renamed

```json
// User expects "MyReport.sql" to become "Prod-1.sql"
// This will NOT happen
```

**Reason**: Only unsaved queries (SQLQueryX.sql) are renamed to preserve file names

### Mistake 3: Wrong wildcard syntax

```json
// WRONG
"server": "PROD*"        // Shell glob syntax
"server": "PROD."        // Regex syntax
"server": "PROD[0-9]"    // Regex syntax

// CORRECT
"server": "PROD%"        // SQL LIKE syntax
"server": "PROD_"        // SQL LIKE syntax (single char)
```

## Next Steps

- [FAQ](FAQ.md) - Frequently asked questions
- [Configuration Guide](Configuration-Guide.md) - Detailed configuration reference
- [Tips & Best Practices](Tips-and-Best-Practices.md) - Optimize your setup
