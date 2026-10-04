# PPT Tools

PowerPoint ribbon add-in (VBA). Repo: https://github.com/brandonkw123/ppt-tools (branch `main`). This folder is the only working copy.

## Layout
- `src/*.bas`: VBA source. Keep them **ASCII only** with CRLF line endings, because non-ASCII characters get garbled when VBA imports them.
- `ribbon/customUI.xml`: the master ribbon definition. `build.ps1 pack` copies it into `deck/customUI/customUI.xml`.
- `deck/`: the unzipped `PPTTools.pptm`. Its macros are compiled into `deck/ppt/vbaProject.bin`, which a script can't edit.
- `build/`: build output (gitignored).

## Build workflow
- **Manual** (README "Rebuilding after a change"): `.\build.ps1 pack`, then import the modules in the VBA editor (Alt+F11), Save As `.ppam`, then `.\build.ps1 unpack`.
- **Automated** (the user prefers this): only works while the user has ticked "Trust access to the VBA project object model" (registry `HKCU\Software\Microsoft\Office\16.0\PowerPoint\Security\AccessVBOM` = 1).
  - **Claude must never change that setting itself.** Ask the user to tick it, and remind them to untick it when finished.
  - With PowerPoint closed, script it over COM:
    1. Open `build\PPTTools.pptm` with no window, then remove and re-import all standard modules from `src\`.
    2. Compile with `VBE.CommandBars.FindControl(1, 578).Execute()`.
    3. Save, then `SaveAs` to a **temporary name** with format 30 (.ppam) and rename it to `PPTTools.ppam`.
       - Saving directly as `PPTTools.ppam` fails while the installed add-in of that name is loaded.
       - `SaveCopyAs` to .ppam always fails.
- **Verify** the built .ppam against `src\` and `ribbon\`. VBA can be extracted from vbaProject.bin with a small Python OLE + MS-OVBA decompressor; oletools isn't installed.
- **Gotcha: VBA re-cases identifiers project-wide**, and the casing persists in vbaProject.bin.
  - Never name a variable after an object property (e.g. `path` turned `pres.Path` into `pres.path`).
  - If casing gets polluted, restore a clean `deck/ppt/vbaProject.bin` and rebuild.
- **Automated tests:** add a helper macro to a throwaway deck: `Sub T(n): Application.Run "PPTTools.ppam!" & n, Nothing`.
  - Call it via `[System.__ComObject].InvokeMember('Run', ...)`. PowerShell's normal `$app.Run(...)` with arguments throws "type must not be ByRef".
  - Commands that show MsgBox, FileDialog or Outlook can't be automated; the user tests those.

## Installed copy
- `%AppData%\Microsoft\AddIns\PPTTools.ppam` is what PowerPoint loads.
  - It's registered at `HKCU\...\PowerPoint\AddIns\PPTTools` (Path=PPTTools.ppam, AutoLoad=1).
  - That folder is a default Trusted Location.
  - Replacing the file is enough to update it, with PowerPoint closed.
- `PPTTools.pptm` in that folder is a legacy copy. The source is now this repo.
- Releases: attach `build\PPTTools.ppam` to a GitHub Release. Users must Unblock the downloaded file (Properties > Unblock).
- Install options:
  - **Manual install via File > Options > Add-ins must always stay available**, because many users are on company-managed PCs where scripts are blocked.
  - The user also wants an optional one-step installer for personal PCs (decided 2026-10-04).

## Releasing
- GitHub CLI is installed (winget, user scope) at `%LOCALAPPDATA%\Microsoft\WinGet\Packages\GitHub.cli_Microsoft.Winget.Source_8wekyb3d8bbwe\bin\gh.exe` and signed in as brandonkw123. Old shells may not have it on PATH yet.
- Steps:
  1. Add a CHANGELOG entry for the new version.
  2. Build, verify and install the add-in.
  3. Commit and push.
  4. Run `gh release create vX.Y.Z build/PPTTools.ppam --repo brandonkw123/ppt-tools --target main --title "PPT Tools vX.Y.Z" --notes-file <notes>`. The notes are the install steps plus that version's CHANGELOG section; see release v1.1.0 for the format.
- The README's install steps link to the Releases page, so the asset must be named `PPTTools.ppam`.

## Code conventions
- Ribbon callbacks are `Public Sub Name(control As IRibbonControl)`.
  - Keep existing button ids and onAction names stable, because users' Quick Access Toolbar entries reference them.
- Get shapes via `SelectedShapes(min, [max])` in modGlobals: `Set sr = SelectedShapes(2): If sr Is Nothing Then Exit Sub`.
- **User preference:** invalid selection or a no-op means `Beep` (the Windows error sound), never a popup.
  - Popups are only for real information: "save first", the PDF path, the cleanup result.
- Avoid `On Error`, except around a single risky call (see `SaveBackup` in modCleanup).
- Resize through `SetSize` in modSize, which handles Lock Aspect Ratio.
- The cleanup commands save a backup to `%TEMP%\PPT Tools backups` before removing anything.

## Icons (imageMso)
- **Every button must have a unique, at-a-glance icon.** The user puts many of them on the Quick Access Toolbar as icons only.
- An invalid name doesn't raise an error; the button just shows no icon. Not every valid image name is a PowerPoint command name, so `GetLabelMso` isn't a complete check.
- **Preview icons:**
  - `CommandBars.GetImageMso` only works in-process. From a VBA macro in a throwaway deck, use GetObject on `pic.Handle` to reach the 32bpp DIB, then copy `bmBits` to a file.
  - The rows are **top-down**, with straight (not premultiplied) alpha.
  - Name list: bert-toolkit.com/imagemso-list.html.
- Current icons (approved by the user 2026-10-04):

| Group | Button | imageMso |
|---|---|---|
| Position | Swap Positions | ArrowsMore |
| Position | Copy Position | Copy |
| Position | Paste Position | Paste |
| Size | Match Width | SizeToWidest |
| Size | Match Height | SizeToTallest |
| Size | Match Width + Height | SizeToControlHeightAndWidth |
| Size | Copy Size | TableResize |
| Size | Paste Width | ShapeWidth |
| Size | Paste Height | ShapeHeight |
| Size | Paste Width + Height | DiagramScale |
| Stretch to Meet | Stretch Up / Down / Left / Right to Meet | ObjectsAlignTopSmart / BottomSmart / LeftSmart / RightSmart |
| Stretch to Match | Stretch Up / Down / Left / Right to Match | FillUp / FillDown / FillLeft / FillRight |
| Email | Email Selected Slides | SendCopySendToMailRecipient |
| Email | Email Whole Deck | CreateEmail |
| Utilities | Convert to PDF | PublishToPdfOrEdoc |
| Cleanup | Remove All Comments | ReviewDeleteComment |
| Cleanup | Remove All Speaker Notes | TableOfContentsRemove |
| Cleanup | Remove Comments + Notes | ReviewDeleteAllMarkupInPresentation |

## Working with the user
- Keep plans and checkpoints short.
- Explain what will change before touching files.
- Use few checkpoints: review the changes, build/test, before pushing to GitHub, before deleting anything.
- The user wants to do as little as possible by hand: automate everything allowed and ask only for what can't be automated.
