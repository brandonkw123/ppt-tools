# Changelog

## [1.1.0] — 2026-10-03

### Added
- Stretch to Match group: Stretch Up/Down/Left/Right to Match line up the moving edge with the same edge of the reference shape (the existing Stretch commands are now named "to Meet")
- Cleanup group: Remove All Comments, Remove All Speaker Notes, Remove Comments + Notes (whole deck; a backup copy of the deck is saved to `%TEMP%\PPT Tools backups` first)
- Hover descriptions on the Stretch buttons
- `build.ps1` and `deck/` folder so the add-in can be rebuilt from this repo
- `LICENSE` file (MIT)

### Fixed
- Blank icons on Swap Positions, Match/Paste Width, Match/Paste Height, Match/Paste Width + Height, Email Selected Slides, Email Whole Deck and Convert to PDF
- Email Selected Slides dropped the first selected slide and lost the deck's theme; it now sends a copy of the deck with only the selected slides
- Match, Paste and Stretch commands also changed the other dimension on shapes with Lock Aspect Ratio turned on (e.g. pictures)
- Stretch commands showed a VBA error when a shape was on the wrong side of the reference shape; those shapes are now skipped and the error sound plays
- Paste Position / Paste Size did nothing, with no feedback, when nothing had been copied; the error sound now plays
- `.bas` files now import without garbled characters in their header comments

### Changed
- Convert to PDF now asks where to save, and can no longer overwrite the original deck
- Every button now has its own distinct icon, so they're recognizable at a glance on the Quick Access Toolbar
- "Please select…" popups replaced by the Windows error sound; a cursor inside a text box now counts as selecting that shape

## [0.1.0] — Initial Release

### Added
- Swap Positions — swap the position of two selected shapes
- Copy Position / Paste Position
- Match Width, Match Height, Match Width + Height
- Copy Size / Paste Width / Paste Height / Paste Width + Height
- Stretch Up, Stretch Down, Stretch Left, Stretch Right
- Email Selected Slides — export selected slides to new .pptx and attach to Outlook draft
- Email Whole Deck — attach current presentation to Outlook draft
- Convert to PDF — one-click export to PDF in same folder as presentation

### Planned
- Position + Size Wizard (sidebar UI)
- Select Similar (select all shapes matching properties of selected shape)
