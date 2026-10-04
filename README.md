# PPT Tools — PowerPoint Productivity Add-in

A custom PowerPoint ribbon add-in that adds fast, keyboard-friendly commands for shape manipulation, sizing, positioning, and file utilities. Inspired by enterprise consulting toolbars.

---

## Commands

### Position
| Command | Behavior |
|---|---|
| Swap Positions | Swaps the position of exactly two selected shapes |
| Copy Position | Stores the position of the selected shape |
| Paste Position | Applies the stored position to all selected shapes |

### Size
| Command | Behavior |
|---|---|
| Match Width to Selected | Resizes all selected shapes to the width of the first selected |
| Match Height to Selected | Resizes all selected shapes to the height of the first selected |
| Match Width & Height to Selected | Resizes all selected shapes to the width and height of the first selected |
| Copy Height & Width | Stores the width and height of the selected shape |
| Paste Width | Applies stored width to all selected shapes |
| Paste Height | Applies stored height to all selected shapes |
| Paste Height & Width | Applies stored width and height to all selected shapes |

Size and Stretch commands change only the dimension you asked for, even on pictures and other shapes with Lock Aspect Ratio turned on.

### Stretch
Each stretch moves one edge of every selected shape to a line on the reference shape; the opposite edge stays put.

| Command | Behavior |
|---|---|
| Stretch Up to Meet | Pulls the top edge up to the reference's **bottom** edge, so the shapes touch |
| Stretch Down to Meet | Pulls the bottom edge down to the reference's **top** edge, so the shapes touch |
| Stretch Left to Meet | Pulls the left edge to the reference's **right** edge, so the shapes touch |
| Stretch Right to Meet | Pulls the right edge to the reference's **left** edge, so the shapes touch |
| Stretch Up to Match | Pulls the top edge to the reference's **top** edge, so the top edges line up |
| Stretch Down to Match | Pulls the bottom edge to the reference's **bottom** edge, so the bottom edges line up |
| Stretch Left to Match | Pulls the left edge to the reference's **left** edge, so the left edges line up |
| Stretch Right to Match | Pulls the right edge to the reference's **right** edge, so the right edges line up |

If the target line is on the wrong side of a shape's fixed edge (the shape would end up with zero or negative size), that shape is skipped and the Windows error sound plays.

### Utilities
| Command | Behavior |
|---|---|
| Email Selected Slides | Saves a copy of the deck containing only the slides selected in the slide panel (original formatting kept) and opens an Outlook draft with it attached |
| Email Whole Deck | Saves the current presentation and opens an Outlook draft with it attached |
| Convert to PDF | Export to PDF — prompts you to choose the save location |

### Cleanup
All cleanup commands act on every slide in the deck. Before removing anything they save a backup copy of the deck to `%TEMP%\PPT Tools backups`.

| Command | Behavior |
|---|---|
| Remove All Comments | Deletes every comment (and its replies) in the presentation |
| Remove All Speaker Notes | Clears the speaker notes text on every slide |
| Remove Comments + Notes | Both of the above in one step |

---

## Selection Behavior

For commands that use a **reference shape** (Match, Stretch), the reference is the selected shape furthest back in the stacking order (z-order), not the one you clicked first. Use **Send to Back** on a shape to make it the reference.

If a command can't run on the current selection (wrong number of shapes, nothing copied yet, and so on), the Windows error sound plays instead of a popup.

---

# Quick Install (for users)

If you just want to use the tool, this is all you need. No coding required. **Close PowerPoint first.**

## Option A — Installer (personal PCs)

1. From the [Releases page](https://github.com/brandonkw123/ppt-tools/releases), download `PPT-Tools.zip` from the latest release
2. Right click it > **Extract All** > **Extract**
3. In the extracted folder, double-click **Install PPT Tools.cmd**
   - If Windows shows **"Windows protected your PC"**, click **More info** > **Run anyway**. This appears because the installer isn't code-signed; you can open the `.cmd` in Notepad to see exactly what it does.
4. Open PowerPoint — the **PPT Tools** tab is there

To update, run the newer installer the same way. To remove, run **Uninstall PPT Tools.cmd**.

## Option B — Manual install (company-managed PCs, or if the installer is blocked)

### 1 — Download the add-in

From the [Releases page](https://github.com/brandonkw123/ppt-tools/releases), download `PPTTools.ppam` from the latest release.

### 2 — Unblock it and place it in the PowerPoint AddIns folder

1. Right click the downloaded `PPTTools.ppam` > **Properties**, tick **Unblock** at the bottom, and click **OK**. Windows marks downloaded files, and Office blocks macros in them until they're unblocked. If there's no Unblock checkbox, the file is already unblocked.
2. Open File Explorer
3. In the address bar, type `%AppData%\Microsoft\AddIns` and press Enter
4. Move `PPTTools.ppam` into this folder

### 3 — Register the add-in

1. Open PowerPoint
2. File > Options > Add-ins
3. At the bottom, set the **Manage** dropdown to **PowerPoint Add-ins** and click **Go**
4. Click **Add**, select `PPTTools.ppam`, and confirm
5. If prompted about macros, choose **Enable Macros**

### 4 — If the buttons don't run

If the tab appears but clicking buttons does nothing:

1. Check that `PPTTools.ppam` is in `%AppData%\Microsoft\AddIns` (PowerPoint trusts add-ins in that folder by default) and, for a manual install, that you unblocked it
2. Close and reopen PowerPoint
3. On a company-managed PC, IT policy may block macros entirely. If so, ask your IT team; don't lower your macro security settings

The **PPT Tools** tab will now appear on every presentation you open.

### Optional — Add buttons to the Quick Access Toolbar

Right click any button in the PPT Tools tab and select **Add to Quick Access Toolbar** for one-click access.

---

# Building from Source (for developers)

Only needed if you want to modify the commands or rebuild the add-in yourself.

### Requirements
- Microsoft PowerPoint (Windows)
- Microsoft Outlook (for email commands)

### Repo layout
| Path | What it is |
|---|---|
| `src/*.bas` | VBA source for every command (plain ASCII, imported into the deck) |
| `ribbon/customUI.xml` | The ribbon tab definition: buttons, labels, icons |
| `deck/` | The unzipped `PPTTools.pptm` the add-in is built from, including its compiled macros |
| `build.ps1` | Packs `deck/` into a `.pptm` and unpacks it back |
| `build/` | Build output (not committed) |

### Rebuilding after a change
1. Edit `src/*.bas` and/or `ribbon/customUI.xml`
2. In PowerShell, from the repo folder: `.\build.ps1 pack` (copies the ribbon XML into the deck and creates `build\PPTTools.pptm`)
3. **If you changed any `.bas` file:** open `build\PPTTools.pptm`, press **Alt+F11**, right click each changed module > **Remove** > **No**, then **File > Import File** and pick the new `.bas` from `src\`. Press **Ctrl+S**.
4. File > Save As > **PowerPoint Add-in (*.ppam)** > save as `build\PPTTools.ppam`, then close PowerPoint
5. `.\build.ps1 unpack` (copies the saved `.pptm` back into `deck/` so the stored macros stay current)
6. To install: copy `build\PPTTools.ppam` into `%AppData%\Microsoft\AddIns`, replacing the old one
7. To release: `.\build.ps1 release` (creates `build\PPT-Tools.zip` with the add-in and the installer scripts from `installer\`), then attach both `build\PPTTools.ppam` and `build\PPT-Tools.zip` to a new GitHub Release

Only use built-in icon names (`imageMso`) that exist in PowerPoint. An unknown name doesn't cause an error; the button just shows no icon.

---

## License

MIT License — see `LICENSE` for details.
