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

### Stretch
| Command | Behavior |
|---|---|
| Stretch Up | Pulls the top edge of all selected shapes up to meet the bottom edge of the first selected |
| Stretch Down | Pulls the bottom edge of all selected shapes down to meet the top edge of the first selected |
| Stretch Left | Pulls the left edge of all selected shapes to meet the right edge of the first selected |
| Stretch Right | Pulls the right edge of all selected shapes to meet the left edge of the first selected |

### Utilities
| Command | Behavior |
|---|---|
| Email Selected Slides | Exports slides selected in the slide panel to a new .pptx and opens an Outlook draft with it attached |
| Email Whole Deck | Saves the current presentation and opens an Outlook draft with it attached |
| Convert to PDF | Export to PDF — prompts you to choose the save location |

---

## Selection Behavior

For commands that reference the **first selected shape** (Match, Stretch), select the reference shape first, then hold **Shift** to add the remaining shapes. PowerPoint passes shapes in z-order, not click order, so the first shape in z-order among your selection acts as the reference.

---

# Quick Install (for users)

If you just want to use the tool, this is all you need. No coding required.

### 1 — Download the add-in

Go to the [Releases page](https://github.com/brandonkw123/ppt-tools/releases) and download `PPTTools.ppam` from the latest release.

### 2 — Place it in the PowerPoint AddIns folder

1. Open File Explorer
2. In the address bar, type `%AppData%\Microsoft\AddIns` and press Enter
3. Move the downloaded `PPTTools.ppam` into this folder

### 3 — Register the add-in

1. Open PowerPoint
2. File > Options > Add-ins
3. At the bottom, set the **Manage** dropdown to **PowerPoint Add-ins** and click **Go**
4. Click **Add**, select `PPTTools.ppam`, and confirm
5. If prompted about macros, choose **Enable Macros**

### 4 — Enable macros (if the buttons don't run)

If the tab appears but clicking buttons does nothing, macros are disabled:

1. File > Options > Trust Center > Trust Center Settings > Macro Settings
2. Select **Enable all macros**
3. Click OK, then close and reopen PowerPoint

The **PPT Tools** tab will now appear on every presentation you open.

### Optional — Add buttons to the Quick Access Toolbar

Right click any button in the PPT Tools tab and select **Add to Quick Access Toolbar** for one-click access.

---

# Building from Source (for developers)

Only needed if you want to modify the commands or rebuild the add-in yourself.

### Requirements
- Microsoft PowerPoint (Windows)
- Microsoft Outlook (for email commands)
- [Office RibbonX Editor](https://github.com/fernandreu/office-ribbonx-editor/releases) (for injecting the ribbon XML)

### 1 — Create the macro-enabled presentation
1. Open PowerPoint, create a new blank presentation
2. File > Save As
3. In the save dialog address bar, type `%AppData%\Microsoft\AddIns` and press Enter
4. Set "Save as type" to **PowerPoint Macro-Enabled Presentation (*.pptm)**
5. Name it `PPTTools` and save

### 2 — Import the VBA modules
1. If the Developer tab isn't visible: File > Options > Customize Ribbon > check **Developer** > OK
2. Developer tab > Visual Basic
3. Right click the `PPTTools` project in the left panel > Import File
4. Import each `.bas` file from the `src/` folder: `modGlobals`, `modPosition`, `modSize`, `modUtilities`
5. Remove the default Module1 if present (right click > Remove > No)
6. Save and close the VBA editor

### 3 — Inject the ribbon XML
1. Close `PPTTools.pptm` in PowerPoint completely
2. Open the Office RibbonX Editor
   - If Windows Defender blocks it: click **More info** > **Run anyway**
3. File > Open, select `PPTTools.pptm`
4. Right click the file in the left panel > **Insert Office 2010 Custom UI Part**
5. Click the new `customUI14.xml` entry
6. Paste in the full contents of `ribbon/customUI.xml`
7. Save and close the editor

### 4 — Verify
1. Open `PPTTools.pptm`, enable macros if prompted
2. Confirm the **PPT Tools** tab appears and the buttons work

### 5 — Convert to .ppam
1. With `PPTTools.pptm` open, File > Save As
2. Navigate to `%AppData%\Microsoft\AddIns`
3. Set "Save as type" to **PowerPoint Add-in (*.ppam)**
4. Keep the name `PPTTools` and save
5. Keep the `.pptm` — it's your editable source for future changes

### 6 — Register
Follow the Quick Install registration steps above to load the `.ppam`.

---

## License

MIT License — see `LICENSE` for details.
