# Line Indicator Targeting Notes

## What we learned on 2026-04-09

- The visual element currently being selected is:
  - `Microsoft.VisualStudio.Shell.Controls.Indicator`
  - Name: `AccentRect`
  - Typical size seen in logs: `w=3 h=16`
- This is associated with the tab header accent in the document well, not the desired editor-side line indicator.
- Coloring this target can still be useful for a future feature (tab header accent customization).

## Future feature idea

- Reuse the current discovery path for optional tab-header accent coloring.
- Keep this separate from line-indicator coloring.

## Known visual caveat

- Unfocused tabs currently show sharp corners when the header accent is recolored.
- If we ship tab-header accent coloring later, we should account for corner styling on unfocused tabs.

## Current direction for line indicator

- Exclude `AccentRect` and tab/document-well ancestors when selecting the line indicator target.
- Prefer indicator candidates that are:
  - left-edge aligned
  - below the top tab strip region
  - under editor/textview/margin/adornment ancestry
- Keep candidate scoring logs enabled until target is stable in SSMS 18.
