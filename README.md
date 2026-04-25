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
| Match Width | Resizes all selected shapes to the width of the first selected |
| Match Height | Resizes all selected shapes to the height of the first selected |
| Match Width + Height | Resizes all selected shapes to the width and height of the first selected |
| Copy Size | Stores the width and height of the selected shape |
| Paste Width | Applies stored width to all selected shapes |
| Paste Height | Applies stored height to all selected shapes |
| Paste Width + Height | Applies stored width and height to all selected shapes |

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
| Convert to PDF | One-click export to PDF, saved in the same folder as the presentation |

---

## Requirements

- Microsoft PowerPoint (Windows)
- Microsoft Outlook (for email commands)
- Macros must be enabled

---

## Installation

### 1 — Download the files

Download or clone this repository.

### 2 — Import the VBA Modules

1. Open PowerPoint and create a new blank presentation
2. Save it as a **PowerPoint Macro-Enabled Presentation (.pptm)** named `PPTTools.pptm`
3. Save it to this location: `C:\Users\[You]\AppData\Roaming\Microsoft\AddIns\`
4. Open the VBA editor via Developer tab > Visual Basic
5. Right click your project in the left panel > Import File
6. Import each `.bas` file from the `src/` folder in this order:
   - `modGlobals.bas`
   - `modPosition.bas`
   - `modSize.bas`
   - `modUtilities.bas`
7. Delete the default Module1 if present (right click > Remove > No)
8. Save and close the VBA editor

### 3 — Inject the Ribbon XML

1. Close `PPTTools.pptm` in PowerPoint completely
2. Rename `PPTTools.pptm` to `PPTTools.zip`
3. Extract all contents to a folder called `PPTTools` on your Desktop
4. Inside that folder, create a new folder called `customUI`
5. Copy `ribbon/customUI.xml` into that `customUI` folder
6. Open the `_rels` folder and open `.rels` in Notepad
7. Add this line before the closing `</Relationships>` tag:
```xml
<Relationship Id="rId5" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="customUI/customUI.xml"/>
```
8. Save and close `.rels`
9. Select everything **inside** the `PPTTools` folder (not the folder itself)
10. Right click > Send to > Compressed (zipped) folder
11. Name it `PPTTools.zip`, then rename it to `PPTTools.pptm`
12. Move it back to `C:\Users\[You]\AppData\Roaming\Microsoft\AddIns\`
13. Open it in PowerPoint — the **PPT Tools** tab should appear

### 4 — Enable Macros

Go to File > Options > Trust Center > Trust Center Settings > Macro Settings and select **Enable all macros**.

### 5 — Add to Quick Access Toolbar (Optional)

Right click any button in the PPT Tools tab and select **Add to Quick Access Toolbar** for one-click access.

### 6 — Convert to .ppam (Optional, Recommended)

Once everything is working, convert to a proper add-in so it loads automatically:

1. Close `PPTTools.pptm` in PowerPoint
2. Rename to `PPTTools.zip` and extract
3. Open `[Content_Types].xml` in Notepad
4. Change every instance of `pptm` to `ppam`
5. Save, rezip, rename to `PPTTools.ppam`
6. In PowerPoint: File > Options > Add-ins > Manage: PowerPoint Add-ins > Go > Add
7. Browse to and enable `PPTTools.ppam`

---

## Selection Behavior

For commands that reference the **first selected shape** (Match, Stretch), select the reference shape first, then hold **Shift** to add the remaining shapes. PowerPoint passes shapes in z-order, not click order, so the first shape in z-order among your selection acts as the reference.

---

## License

MIT License — see `LICENSE` for details.
