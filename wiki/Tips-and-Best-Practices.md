# Tips and Best Practices

Optimize your SSMS EnvTabs configuration for maximum productivity.

## Configuration Best Practices

### 1. Use Priority Strategically

**Rule**: Lower priority numbers = checked first

**Best Practice**: Use increments of 10 (10, 20, 30, ...) to allow inserting rules later

```json
{
  "groups": [
    {
      "groupName": "Critical-Prod",
      "server": "PROD-CORE-%",
      "priority": 10,
      "colorIndex": 3
    },
    // Can insert priority 15 here later if needed
    {
      "groupName": "Prod",
      "server": "%PROD%",
      "priority": 20,
      "colorIndex": 7
    },
    // Can insert priority 25 here later
    {
      "groupName": "QA",
      "server": "%QA%",
      "priority": 30,
      "colorIndex": 1
    }
  ]
}
```

### 2. Specific Rules First, General Rules Last

```json
{
  "groups": [
    {
      "groupName": "Prod-Main",
      "server": "PROD-SQL-01",
      "database": "CustomerDB",
      "priority": 10,
      "colorIndex": 3,
      "comment": "Most specific"
    },
    {
      "groupName": "Prod",
      "server": "%PROD%",
      "database": "%",
      "priority": 20,
      "colorIndex": 3,
      "comment": "Broader production"
    },
    {
      "groupName": "Other",
      "server": "%",
      "database": "%",
      "priority": 100,
      "colorIndex": 8,
      "comment": "Catch-all"
    }
  ]
}
```

### 3. Keep Group Names Short

**Why**: Tab space is limited in SSMS

```json
// GOOD
"groupName": "Prod"         // Short and clear
"groupName": "QA"
"groupName": "Dev"

// LESS IDEAL
"groupName": "Production-Environment"    // Too long
"groupName": "QA-Testing-Server"
```

**Recommendation**: 4-8 characters maximum

### 4. Document Your Configuration

Add comments to explain your rules (use a `comment` field that the extension ignores):

```json
{
  "groupName": "Prod",
  "server": "%PROD%",
  "database": "%",
  "priority": 10,
  "colorIndex": 3,
  "comment": "All production servers - USE CAUTION"
}
```

Or maintain a separate documentation file.

### 5. Use Consistent Naming Conventions

**For Group Names**:
- Use consistent capitalization: "Prod", "QA", "Dev" (not "prod", "qa", "DEV")
- Use abbreviations consistently
- Match your team's terminology

**For Patterns**:
- If your servers use hyphens, match that: `"PROD-%"`
- If they use underscores, match that: `"PROD_%"`

## Color Strategy

### 1. Reserve Red Tones for Production

**Burgundy (3)** and **Pumpkin (7)** are attention-grabbing:

```json
{
  "groups": [
    {
      "groupName": "Prod",
      "server": "%PROD%",
      "colorIndex": 3,
      "comment": "Burgundy - STOP AND THINK"
    }
  ]
}
```

This creates a visual warning before running queries on production.

### 2. Use Color Families for Related Environments

Group related environments with similar colors:

```json
{
  "groups": [
    {
      "groupName": "Prod-US",
      "server": "%US-PROD%",
      "colorIndex": 3,
      "comment": "Burgundy for all production"
    },
    {
      "groupName": "Prod-EU",
      "server": "%EU-PROD%",
      "colorIndex": 7,
      "comment": "Pumpkin also production (similar red tone)"
    }
  ]
}
```

### 3. Use Neutral Colors for Read-Only

**Gray (8)** or **Brown (5)** work well for reporting/read-only databases:

```json
{
  "groupName": "Reports",
  "server": "%",
  "database": "%Reports%",
  "colorIndex": 8,
  "comment": "Gray - read-only, low risk"
}
```

### 4. Use Green/Blue for Development

**Green (4)**, **Cyan (2)**, **Blue (14)** are calming and indicate "safe to experiment":

```json
{
  "groupName": "Dev",
  "server": "%DEV%",
  "colorIndex": 4,
  "comment": "Green - safe environment"
}
```

### 5. Document Your Color Scheme

Create a team reference:

```
Production:  Burgundy (3), Pumpkin (7)
QA/Staging:  Gold (1), Pink (15)
Development: Green (4), Cyan (2), Blue (14)
Test:        Volt (9), Mint (12)
Read-Only:   Gray (8), Brown (5)
Special:     Magenta (11), Royal Blue (6)
```

## Pattern Matching Tips

### 1. Test Patterns Before Deploying

Use a test rule at priority 1 to verify your pattern:

```json
{
  "groupName": "TEST",
  "server": "YOUR-NEW-PATTERN",
  "database": "%",
  "priority": 1,
  "colorIndex": 15,
  "comment": "TEMPORARY - Remove after testing"
}
```

### 2. Be Careful with Broad Patterns

```json
// RISKY - Might match more than intended
"server": "%PROD%"    // Matches PREPROD, PROD, PRODUCTION

// SAFER - More specific
"server": "PROD-%"    // Only matches PROD-xxx
```

### 3. Use the Underscore Wildcard for Format Matching

For servers with consistent naming:

```json
{
  "server": "SQL-0_",
  "comment": "Matches SQL-01 through SQL-09 only"
}
```

### 4. Consider Special Characters

Some server names have spaces, hyphens, or underscores:

```json
"server": "% PROD %"      // Space before and after
"server": "%PROD_%"       // Underscore
"server": "%-PROD-%"      // Hyphens
```

## Workflow Optimization

### 1. Use Auto-Configuration for New Environments

When exploring new environments, enable auto-configuration:

```json
"settings": {
  "autoConfigure": "server db",
  "enableConfigurePrompt": true
}
```

Then refine the auto-generated rules later.

### 2. Disable Auto-Rename for Saved Files

This is the default behavior (saved files keep their names):

```json
"settings": {
  "enableAutoRename": true   // Only renames SQLQueryX.sql
}
```

**Why**: Preserves your carefully named script files.

### 3. Create a "Catch-All" Rule

Always have a final catch-all rule to see what's not matching:

```json
{
  "groupName": "Other",
  "server": "%",
  "database": "%",
  "priority": 999,
  "colorIndex": 8,
  "comment": "Catch anything not matched above"
}
```

If you see "Other-1" tabs, you know those connections need rules.

### 4. Group By Project, Not Just Environment

For complex systems:

```json
{
  "groups": [
    {
      "groupName": "ERP-Prod",
      "server": "%PROD%",
      "database": "%ERP%",
      "priority": 10,
      "colorIndex": 3
    },
    {
      "groupName": "Web-Prod",
      "server": "%PROD%",
      "database": "%Website%",
      "priority": 11,
      "colorIndex": 7
    }
  ]
}
```

### 5. Create Environment-Specific Configurations

For different roles:

**DBA Configuration** (focused on servers):
```json
{
  "groups": [
    {"groupName": "Prod", "server": "%PROD%", "priority": 10},
    {"groupName": "QA", "server": "%QA%", "priority": 20},
    {"groupName": "Dev", "server": "%DEV%", "priority": 30}
  ]
}
```

**Developer Configuration** (focused on projects):
```json
{
  "groups": [
    {"groupName": "API", "database": "%API%", "priority": 10},
    {"groupName": "Web", "database": "%Web%", "priority": 20},
    {"groupName": "Mobile", "database": "%Mobile%", "priority": 30}
  ]
}
```

## Team Collaboration

### 1. Share Configuration Files

Store configuration in version control:

```bash
# Save your config to a shared location
copy "%USERPROFILE%\Documents\SSMS EnvTabs\TabGroupConfig.json" "C:\TeamConfigs\"
```

Team members can then copy it to their own Documents folder.

### 2. Use Descriptive Group Names

Make it obvious what each group represents:

```json
// GOOD
"groupName": "Prod-US-East"

// LESS CLEAR
"groupName": "Grp1"
```

### 3. Document Your Color Choices

Create a team wiki or README explaining your color scheme.

### 4. Standardize Across the Team

Have all team members use the same:
- Group names
- Color scheme
- Priority strategy

This helps when discussing issues or sharing screenshots.

## Maintenance Tips

### 1. Review Configuration Regularly

Every few months:
- Remove obsolete rules for decommissioned servers
- Add rules for new environments
- Optimize patterns that are too broad or too narrow

### 2. Backup Your Configuration

Before major changes:

```bash
copy "%USERPROFILE%\Documents\SSMS EnvTabs\TabGroupConfig.json" "%USERPROFILE%\Documents\SSMS EnvTabs\TabGroupConfig-backup-2024-01-15.json"
```

### 3. Keep It Simple

Don't over-engineer your configuration:
- Fewer rules are easier to maintain
- Simpler patterns are easier to understand
- Not every scenario needs a unique color

### 4. Test Changes Incrementally

Don't rewrite your entire configuration at once:
1. Make one small change
2. Save and test
3. Verify it works as expected
4. Move to the next change

### 5. Monitor for Issues

Check the log file occasionally:

```
%LocalAppData%\SSMS EnvTabs\runtime.log
```

Look for parsing errors or unexpected behavior.

## Performance Considerations

### 1. Avoid Excessive Rules

**Good**: 5-10 rules
**Acceptable**: 10-20 rules
**Too many**: 50+ rules (may slow down tab switching)

### 2. Use Efficient Patterns

```json
// EFFICIENT
"server": "PROD-%"        // Simple prefix match

// LESS EFFICIENT
"server": "%-%-%-%-%"     // Complex pattern with many wildcards
```

### 3. Order Rules by Likelihood

Put frequently-matched rules first (lower priority):

```json
{
  "groups": [
    {
      "groupName": "Dev",
      "server": "%DEV%",
      "priority": 10,
      "comment": "Most commonly used - check first"
    },
    {
      "groupName": "Prod",
      "server": "%PROD%",
      "priority": 20,
      "comment": "Less frequently used"
    }
  ]
}
```

## Security Best Practices

### 1. Use Tab Coloring as a Safety Reminder

Bright colors for production = visual reminder to double-check queries:

```json
{
  "groupName": "Prod",
  "server": "%PROD%",
  "colorIndex": 3,
  "comment": "Burgundy = DANGER"
}
```

### 2. Create Specific Rules for Critical Databases

```json
{
  "groupName": "CRITICAL",
  "server": "PROD-SQL-01",
  "database": "CustomerData",
  "priority": 1,
  "colorIndex": 3,
  "comment": "Highest priority - most visible"
}
```

### 3. Don't Rely Solely on Colors

Tab colors are a visual aid, not a security control:
- Still check your connection before running queries
- Use SQL Server permissions properly
- Consider read-only connections for production

## Advanced Tips

### 1. Use JSON Comments Workaround

Standard JSON doesn't allow comments, but you can use a "comment" field:

```json
{
  "groupName": "Prod",
  "server": "%PROD%",
  "priority": 10,
  "colorIndex": 3,
  "comment": "All production servers - matches PROD-SQL-01, PROD-SQL-02, etc."
}
```

The extension ignores unknown fields.

### 2. Create Templates for Common Scenarios

Keep template configurations for different needs:

**Template: Simple Three-Tier**
```json
{
  "groups": [
    {"groupName": "Prod", "server": "%PROD%", "priority": 10, "colorIndex": 3},
    {"groupName": "QA", "server": "%QA%", "priority": 20, "colorIndex": 1},
    {"groupName": "Dev", "server": "%DEV%", "priority": 30, "colorIndex": 4}
  ]
}
```

**Template: Project-Based**
```json
{
  "groups": [
    {"groupName": "ProjectA", "database": "%ProjectA%", "priority": 10, "colorIndex": 14},
    {"groupName": "ProjectB", "database": "%ProjectB%", "priority": 20, "colorIndex": 4}
  ]
}
```

### 3. Use Priority Ranges for Categories

Reserve priority ranges for different types of rules:
- 1-19: Critical/specific rules
- 20-49: Environment-based rules
- 50-79: Project-based rules
- 80-99: Database-specific rules
- 100+: Catch-all/fallback rules

## Common Anti-Patterns to Avoid

### ❌ Don't: Use Random Priorities

```json
{
  "groups": [
    {"priority": 5},
    {"priority": 47},
    {"priority": 3},
    {"priority": 98}
  ]
}
```

### ✅ Do: Use Consistent Priority Increments

```json
{
  "groups": [
    {"priority": 10},
    {"priority": 20},
    {"priority": 30},
    {"priority": 40}
  ]
}
```

### ❌ Don't: Use Too Many Similar Colors

Avoid using colors that look too similar:

```json
{"colorIndex": 2},   // Cyan
{"colorIndex": 10},  // Teal
{"colorIndex": 14}   // Blue
```

### ✅ Do: Use Visually Distinct Colors

```json
{"colorIndex": 3},   // Burgundy
{"colorIndex": 4},   // Green
{"colorIndex": 14}   // Blue
```

### ❌ Don't: Make Patterns Too Specific

```json
"server": "PROD-SQL-01-SERVER-INSTANCE-1"   // Too specific
```

### ✅ Do: Use Reasonable Patterns

```json
"server": "PROD-SQL-01"   // Appropriately specific
```

## Quick Reference Checklist

Before deploying a configuration:

- [ ] Rules are ordered by priority (lowest first)
- [ ] Specific patterns have lower priority numbers
- [ ] Catch-all rule exists at the end (high priority number)
- [ ] Production uses red tones (Burgundy/Pumpkin)
- [ ] Development uses green/blue tones
- [ ] Group names are short (4-8 characters)
- [ ] Patterns use SQL LIKE syntax (%, _)
- [ ] JSON is valid (no syntax errors)
- [ ] Configuration has been tested
- [ ] Configuration is backed up

## Next Steps

- [Configuration Guide](Configuration-Guide.md) - Complete reference
- [Wildcard Patterns](Wildcard-Patterns.md) - Pattern matching guide
- [Color Reference](Color-Reference.md) - All available colors
- [Troubleshooting](Troubleshooting.md) - Fix common issues
