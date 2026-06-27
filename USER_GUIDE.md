# UpslideClone — User Guide

> A practical, button-by-button guide to every **Upslide** feature across Excel, PowerPoint
> and Word — what each ribbon button does and how to use it.

For install/build, see the [README](README.md).

## Before you start
1. Install once: `installer\install.ps1` (no admin needed).
2. Open the app with the cold-start launcher (`tools\Launch-Excel.cmd`, etc.) and open a
   document — the **Upslide** ribbon tab appears.
3. If the tab is missing: *File ▸ Options ▸ Add-ins ▸ Manage: COM Add-ins ▸ Go…* → tick
   **UpslideClone** (or re-enable from *Disabled Items*). A clean restart usually fixes it.
4. Every command logs to `%APPDATA%\UpslideClone\logs` — handy if something errors.

---

## Excel — the Upslide tab

### Formatting
| Button | What it does | How to use |
|---|---|---|
| **Smart Format** | Branded table: green header, bold result rows (EBITDA, Gross Margin, 毛利…), number/percent formats, borders. | Select the whole table (incl. header) → click. |
| **Undo Formatting** | Reverts the last Smart Format cell-by-cell. | Click after a Smart Format you want to undo. |
| **Clear Formatting** | Strips all formatting (keeps values). | Select the range → click. |
| **Toggles ▸** | Flip one dimension on/off: Title / Result / Number / Percent formats. | Select the table → pick a toggle. |

> Formatting is driven by `assets\theme.json` (and **Settings**). Edit it to change colours,
> fonts or number formats without rebuilding.

### Charts
| Button | What it does | How to use |
|---|---|---|
| **Waterfall** | Floating-bar bridge: dark-green anchors, green increases, red decreases, data labels. | Select a two-column label/value bridge (e.g. 45 → 32 → 53) → click. |
| **Stacked Waterfall** | Multi-series bridge (e.g. FR/UK/DE). | Select header + label column + ≥2 category columns → click. |
| **Display CAGR** | Adds a CAGR % arrow between the first and last points of the selected chart. | Click a chart, then this button. |

### Modelling
| Button | What it does | How to use |
|---|---|---|
| **Autocolor** | Blue = hardcoded inputs, black = formulas, green = links/external refs. | Select a model range → click. |
| **IFERROR** | Wraps selected formulas in `IFERROR(…,"")`. | Select formula cells → click. |
| **Fast Fill Right / Down** | Structure-aware fill of the top-left formula across the selection ($-anchors respected). | Put a formula in the corner, select the fill range → click. |
| **Clean & Prepare** | Removes broken defined names (#REF!) before sharing. | Click; a report lists what was removed. |
| **Smart Print** | Standardised print setup (landscape, fit-to-width) + preview. | Click. |

### Export (the core)
| Button | What it does | How to use |
|---|---|---|
| **Export to PowerPoint** | Pastes the selection into PowerPoint as a **linked picture**. | **Save the workbook first.** Select a range → click. |
| **Export as Table** | Exports as a native, editable, linked PowerPoint table. | Select a range → click. |
| **Export to Word** | Inserts a linked picture at the Word cursor. | Open Word, select an Excel range → click. |
| **Advanced Export** | Batch many ranges → many slides from a mapping table (`Range \| Slide \| Type \| Placeholder`). | Build the mapping table, select it → click. |
| **Highlight Linked** | Flags every range in the workbook that has been linked out. | Click after exporting. |

### Library / Settings
| Button | What it does |
|---|---|
| **Save to Library / Insert from Library** | Save a selection as a reusable snippet; insert it elsewhere. |
| **Settings** | Edit the branding theme path + keyboard shortcuts (persisted per-user). |

---

## PowerPoint — the Upslide tab

### Links
| Button | What it does | How to use |
|---|---|---|
| **Link Manager** | Side pane listing every linked object: slide, type, source file, sheet!range, last refresh, status. | Click to show/hide. |
| **Refresh All** | Re-renders every linked object from its Excel source, preserving position/size. | Click after changing the Excel data. |
| **In the pane:** Reload · **Check Sources** (flags *Changed* via content hash) · Refresh Selected · **Change Source…** (repoint to a moved/renamed file) · Go To. |

### Sizing Guide
| Button | What it does |
|---|---|
| **Insert Placeholder** | Drops a tagged placeholder; exports that name it (Advanced Export's *Placeholder* column) snap to its exact size/position. |

### Design toolkit
| Group | Buttons |
|---|---|
| **Align & Size** | Smart Align (6 modes), Distribute Across/Down, Same Size |
| **Arrange** | Bring to Front, Send to Back, Group, Ungroup |
| **Format & Select** | Format Shapes (brand theme), Select Similar, Smart Painter |
| **Slides** | Table of Contents, Slide Check, Outline pane |
| **Content Library** | Save to Library / Insert from Library (stored under Documents\UpslideClone) |
| **References** | Footnote (auto-numbered), Cross-reference (survives re-ordering), Refresh Refs |

Select the shapes/slides, then click. Align/Distribute need 2–3+ shapes; Smart Painter
copies the first shape's format to the rest.

---

## Word — the Upslide tab

| Button | What it does |
|---|---|
| **Link Manager** | Lists Excel→Word linked objects with source/status; Refresh / Change Source / Go To. |
| **Refresh All** | Re-renders linked pictures/tables from Excel (anchored by bookmarks). |
| **Format Table** | Brand the table the cursor is in: green header row (white bold), brand font, single borders. |
| **Brand Font** | Apply the brand font to the selected text. |

---

## The end-to-end workflow (why this tool exists)
1. In **Excel**, format a table (**Smart Format**) and build a chart (**Waterfall**).
2. **Export to PowerPoint** → a linked object lands on a slide.
3. Change a number in Excel.
4. In **PowerPoint**, open **Link Manager ▸ Refresh All** → the slide updates in place.
5. Moved the workbook? **Change Source…** → repoint → Refresh.

That live, auditable link — Excel model → deck — is the whole point.

---

## Troubleshooting
| Symptom | Fix |
|---|---|
| No Upslide tab | Cold-start via `tools\Launch-*.cmd`; or COM Add-ins ▸ re-enable. After a rebuild, re-run `install.ps1`. |
| A button shows a warning box | The message names the issue; the full stack is in `%APPDATA%\UpslideClone\logs`. |
| Export says "save the workbook first" | Links need a saved source file — save the workbook, then export. |
| Refresh says source missing | The workbook moved/renamed → **Change Source…** in the Link Manager. |
| Tab vanished after I reopened the app | Office kept a background process; fully quit, then use the cold-start launcher. |
