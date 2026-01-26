# Color Reference

SSMS EnvTabs supports 16 distinct colors for tab grouping. Choose colors that help you quickly identify different environments.

## Available Colors

| colorIndex | Color Name   | Use Case Suggestions                                    |
|------------|--------------|--------------------------------------------------------|
| 0          | Lavender     | Non-critical, documentation databases                  |
| 1          | Gold         | QA, staging, pre-production environments               |
| 2          | Cyan         | Development, local instances                           |
| 3          | Burgundy     | **Production** (red tone for alertness)                |
| 4          | Green        | Development, testing, safe environments                |
| 5          | Brown        | Legacy systems, archived databases                     |
| 6          | Royal Blue   | Training, demo environments                            |
| 7          | Pumpkin      | **Production** (orange for caution)                    |
| 8          | Gray         | Read-only, reporting databases                         |
| 9          | Volt         | Development, experimental features                     |
| 10         | Teal         | Testing, integration environments                      |
| 11         | Magenta      | Special projects, temporary environments               |
| 12         | Mint         | Sandbox, prototyping environments                      |
| 13         | Dark Brown   | Maintenance, backup servers                            |
| 14         | Blue         | Standard development or corporate databases            |
| 15         | Pink         | User acceptance testing (UAT), review environments     |

## Color Selection Tips

### By Environment Type

**Production Environments**:
- Use attention-grabbing colors: **Burgundy (3)**, **Pumpkin (7)**
- Helps prevent accidental changes to production

**QA/Staging**:
- Use warning colors: **Gold (1)**, **Orange/Pumpkin (7)**
- Indicates "be careful but not critical"

**Development**:
- Use calming colors: **Green (4)**, **Cyan (2)**, **Blue (14)**
- Indicates safe to experiment

**Special/Temporary**:
- Use distinctive colors: **Magenta (11)**, **Pink (15)**, **Volt (9)**
- Easy to spot temporary or one-off connections

**Read-Only/Reports**:
- Use neutral colors: **Gray (8)**, **Brown (5)**
- Indicates data consumption, not modification

### Color Grouping Strategies

#### Strategy 1: Environment-Based (Recommended)

Use one color per environment type:

```json
{
  "groups": [
    {
      "groupName": "Prod",
      "server": "%PROD%",
      "database": "%",
      "colorIndex": 3
    },
    {
      "groupName": "QA",
      "server": "%QA%",
      "database": "%",
      "colorIndex": 1
    },
    {
      "groupName": "Dev",
      "server": "%DEV%",
      "database": "%",
      "colorIndex": 4
    }
  ]
}
```

#### Strategy 2: Project-Based

Use different colors for different projects/systems:

```json
{
  "groups": [
    {
      "groupName": "Finance",
      "server": "%",
      "database": "%Finance%",
      "colorIndex": 11
    },
    {
      "groupName": "HR",
      "server": "%",
      "database": "%HR%",
      "colorIndex": 14
    },
    {
      "groupName": "Sales",
      "server": "%",
      "database": "%Sales%",
      "colorIndex": 4
    }
  ]
}
```

#### Strategy 3: Criticality-Based

Use color intensity to indicate importance:

```json
{
  "groups": [
    {
      "groupName": "Critical",
      "server": "%PROD%",
      "database": "%Critical%",
      "colorIndex": 3
    },
    {
      "groupName": "Standard",
      "server": "%PROD%",
      "database": "%",
      "colorIndex": 7
    },
    {
      "groupName": "Low-Impact",
      "server": "%",
      "database": "%Reports%",
      "colorIndex": 8
    }
  ]
}
```

#### Strategy 4: Geographic

Use colors to represent regions:

```json
{
  "groups": [
    {
      "groupName": "US-East",
      "server": "%US-E%",
      "database": "%",
      "colorIndex": 14
    },
    {
      "groupName": "US-West",
      "server": "%US-W%",
      "database": "%",
      "colorIndex": 2
    },
    {
      "groupName": "EU",
      "server": "%EU%",
      "database": "%",
      "colorIndex": 4
    },
    {
      "groupName": "APAC",
      "server": "%APAC%",
      "database": "%",
      "colorIndex": 1
    }
  ]
}
```

## Configuration Examples

### Example 1: Simple Three-Tier

```json
{
  "groups": [
    {
      "groupName": "Prod",
      "server": "%PROD%",
      "database": "%",
      "priority": 10,
      "colorIndex": 3,
      "comment": "Burgundy for production - be careful!"
    },
    {
      "groupName": "QA",
      "server": "%QA%",
      "database": "%",
      "priority": 20,
      "colorIndex": 1,
      "comment": "Gold for QA - caution required"
    },
    {
      "groupName": "Dev",
      "server": "%",
      "database": "%",
      "priority": 30,
      "colorIndex": 4,
      "comment": "Green for dev - safe to experiment"
    }
  ]
}
```

### Example 2: Multi-Project Color Coding

```json
{
  "groups": [
    {
      "groupName": "Prod-ERP",
      "server": "%PROD%",
      "database": "%ERP%",
      "priority": 5,
      "colorIndex": 3,
      "comment": "Burgundy for critical ERP production"
    },
    {
      "groupName": "Prod-Web",
      "server": "%PROD%",
      "database": "%Website%",
      "priority": 6,
      "colorIndex": 7,
      "comment": "Pumpkin for web production"
    },
    {
      "groupName": "Prod-Reports",
      "server": "%PROD%",
      "database": "%Reports%",
      "priority": 7,
      "colorIndex": 8,
      "comment": "Gray for read-only reports"
    },
    {
      "groupName": "Dev-ERP",
      "server": "%DEV%",
      "database": "%ERP%",
      "priority": 20,
      "colorIndex": 14,
      "comment": "Blue for ERP development"
    },
    {
      "groupName": "Dev-Web",
      "server": "%DEV%",
      "database": "%Website%",
      "priority": 21,
      "colorIndex": 2,
      "comment": "Cyan for web development"
    }
  ]
}
```

### Example 3: Using All 16 Colors

For organizations with many distinct environments:

```json
{
  "groups": [
    {"groupName": "Prod-US", "server": "%US-PROD%", "database": "%", "priority": 10, "colorIndex": 3},
    {"groupName": "Prod-EU", "server": "%EU-PROD%", "database": "%", "priority": 11, "colorIndex": 7},
    {"groupName": "Prod-APAC", "server": "%APAC-PROD%", "database": "%", "priority": 12, "colorIndex": 13},
    {"groupName": "QA-US", "server": "%US-QA%", "database": "%", "priority": 20, "colorIndex": 1},
    {"groupName": "QA-EU", "server": "%EU-QA%", "database": "%", "priority": 21, "colorIndex": 15},
    {"groupName": "Stage-US", "server": "%US-STAGE%", "database": "%", "priority": 30, "colorIndex": 7},
    {"groupName": "Dev-Team1", "server": "%DEV1%", "database": "%", "priority": 40, "colorIndex": 4},
    {"groupName": "Dev-Team2", "server": "%DEV2%", "database": "%", "priority": 41, "colorIndex": 2},
    {"groupName": "Dev-Team3", "server": "%DEV3%", "database": "%", "priority": 42, "colorIndex": 14},
    {"groupName": "Test", "server": "%TEST%", "database": "%", "priority": 50, "colorIndex": 9},
    {"groupName": "UAT", "server": "%UAT%", "database": "%", "priority": 51, "colorIndex": 15},
    {"groupName": "Demo", "server": "%DEMO%", "database": "%", "priority": 60, "colorIndex": 6},
    {"groupName": "Training", "server": "%TRAIN%", "database": "%", "priority": 61, "colorIndex": 0},
    {"groupName": "Sandbox", "server": "%SANDBOX%", "database": "%", "priority": 70, "colorIndex": 12},
    {"groupName": "Archive", "server": "%ARCHIVE%", "database": "%", "priority": 80, "colorIndex": 5},
    {"groupName": "Local", "server": "localhost", "database": "%", "priority": 90, "colorIndex": 10}
  ]
}
```

## Best Practices

### Do's

✓ **Use consistent colors** across similar environments
  - All production = red tones (Burgundy, Pumpkin)
  - All development = green/blue tones

✓ **Reserve Burgundy (3) for critical production** databases
  - Most attention-grabbing color
  - Universal "stop and think" signal

✓ **Use neutral colors** (Gray, Brown) for read-only databases
  - Indicates less risk

✓ **Document your color choices** in comments or a shared document
  - Helps team members understand the system

✓ **Test colors** with your actual server names
  - Some colors may be more visible than others on your monitor

### Don'ts

✗ **Don't use random colors** without a system
  - Makes it harder to learn and remember

✗ **Don't use similar colors** for different criticality levels
  - Can lead to confusion and mistakes

✗ **Don't change colors frequently**
  - Team members need time to build muscle memory

✗ **Don't use colorIndex values** outside 0-15 range
  - Will cause errors

## Color Accessibility

Consider team members with color vision deficiencies:

- **Use naming conventions** in addition to colors
  - The `groupName` field helps when colors are hard to distinguish
- **Combine colors with priority**
  - Critical environments should also be at top of the list
- **Test your configuration** with color-blind simulators if needed

## Next Steps

- [Configuration Guide](Configuration-Guide.md) - Learn how to configure tab groups
- [Tips & Best Practices](Tips-and-Best-Practices.md) - Optimize your color strategy
