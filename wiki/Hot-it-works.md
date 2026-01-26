# Hot it works

This page explains how SSMS EnvTabs reuses SSMS's built-in **Color Tabs by Regex** mechanism, and how it uses a **salt** to hit the color you request in your rules.

## 1) EnvTabs generates regex rules for SSMS

EnvTabs does **not** paint tabs directly. Instead, it writes regex rules into the same `ColorByRegexConfig.txt` file that SSMS already uses for Color Tabs by Regex. The extension builds a generated block (between `BEGIN`/`END` markers) containing regex lines that map the open documents in each EnvTabs group to the SSMS Color Tabs system.【F:SSMS EnvTabs/ColorByRegexConfigWriter.cs†L11-L17】【F:SSMS EnvTabs/ColorByRegexConfigWriter.cs†L86-L110】

To keep colors stable even when files move between folders, the regex is built from **file names only**. Each group gets a regex that matches any path ending in one of those filenames:

```
(?:^|[\\/])(?:FileA.sql|FileB.sql)$
```

This comes directly from EnvTabs escaping the filenames and assembling a regex that matches any path ending with one of them.【F:SSMS EnvTabs/ColorByRegexConfigWriter.cs†L72-L83】【F:SSMS EnvTabs/ColorByRegexConfigWriter.cs†L117-L123】

## 2) SSMS picks a color by hashing the regex

SSMS chooses a color index by hashing the full regex line and taking `abs(hash) % 16`. EnvTabs mirrors that behavior via a legacy, stable string hash so it can predict the exact color SSMS will choose.【F:SSMS EnvTabs/ColorByRegexConfigWriter.cs†L125-L143】【F:SSMS EnvTabs/TabGroupColorSolver.cs†L27-L35】【F:SSMS EnvTabs/TabGroupColorSolver.cs†L45-L67】

## 3) The salt nudges the hash without changing the match

If your EnvTabs rule specifies a target `ColorIndex`, EnvTabs appends a **regex comment** to the end of the generated regex:

```
(?#salt:123)
```

Because this is a regex comment, it **does not affect matching**, but it **does change the hash input**. EnvTabs brute-forces a numeric salt until the hash result matches your requested color index, then appends it to the regex.【F:SSMS EnvTabs/ColorByRegexConfigWriter.cs†L125-L150】【F:SSMS EnvTabs/TabGroupColorSolver.cs†L10-L39】

## 4) Putting it together

1. EnvTabs groups tabs by your server/database rule.
2. It writes regex lines into `ColorByRegexConfig.txt` for SSMS to consume.
3. If a color is specified, EnvTabs finds a salt so SSMS’s hash lands on that color.
4. SSMS applies the matching color using its built-in Color Tabs by Regex system.

That’s how EnvTabs piggybacks on SSMS’s existing coloring feature while still giving you deterministic, rule-based colors.
