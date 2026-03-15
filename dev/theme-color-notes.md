# Theme Color Notes (SSMS EnvTabs)

## Summary
The toggle switch ON state now tracks SSMS theme accent by using VS theme resource keys exposed through `EnvironmentColors`.

## Key Finding
In SSMS, common accent keys were often null in this control context:
- `SystemAccentBrushKey`
- `SystemAccentColorKey`
- `AccentBrushKey`
- `AccentColorKey`
- `DocumentWellBorderBrushKey`

The most reliable accent-like key for theme identity was:
- `MainWindowActiveDefaultBorderBrushKey`

This key changed with theme and produced expected hue families (for example green/purple/warm variants depending on theme).

## Implemented Resolver Strategy
In `SettingsToolWindowControl.xaml.cs`, toggle accent resolution now uses a simple strategy:
1. `MainWindowActiveDefaultBorderBrushKey`
2. Fallback to the already-resolved local border brush if key 1 is unavailable.

## Runtime Logging
A concise debug line is kept in runtime logs:
- `SettingsToolWindowControl toggle accent source=<key>, color=#RRGGBB`

This confirms which key is currently driving toggle color for each theme.

## Practical Notes
- `FileTabSelectedBorderBrushKey` usually resolves, but often maps to blue/purple tab-selection color rather than overall theme accent.
- Using `MainWindowActiveDefaultBorderBrushKey` better matches the current theme identity in SSMS.
- Keeping one fallback avoids hard failure if the key is unavailable in another shell/version/context.
