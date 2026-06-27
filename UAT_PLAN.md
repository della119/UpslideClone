# UpslideClone — User Acceptance Test (UAT) Plan

_Manual acceptance for W1–W5, mapped to the functional requirements in
`UPSLIDE_CLONE_Windows_Requirements.md`. Tick each result; log issues with the
message from the on-screen box and `%APPDATA%\UpslideClone\logs`._

## 0. Prerequisites (one time)

> **Do NOT double-click `UpslideClone.sln`** — it may open in *Blend* (a XAML tool
> that can't load Office add-in projects). You don't need any IDE to test; build from
> the command line. To open in real Visual Studio for debugging, launch **Visual
> Studio 2022** first then *File ▸ Open ▸ Solution*.

1. Build from a terminal (no IDE):
   `& "D:\VS2022\MSBuild\Current\Bin\MSBuild.exe" UpslideClone.sln /t:Build /restore`
2. Install the add-ins: `powershell -ExecutionPolicy Bypass -File installer\install.ps1`
3. **Close and reopen** Excel, PowerPoint, Word. Each should show an **Upslide** ribbon tab.
   - If a tab is missing: *File ▸ Options ▸ Add-ins ▸ Manage: COM Add-ins ▸ Go…* → tick **UpslideClone** (or re-enable from Disabled Items).
4. Open `Training Guide Excel.xlsx` and **Save As** a copy to your Desktop
   (links need a saved source file). Use this copy throughout.

Legend: **Steps** → **Expected** → ⬜ Pass / ⬜ Fail.

---

## Phase W1 — Branded formatting & charts

### T-W1.1 Smart Format (FR-F1/F2) — tab "Format tables"
- Steps: select the income-statement table → **Upslide ▸ Smart Format**.
- Expected: dark header row (white bold), result rows (Gross Margin, EBITDA) bold +
  light fill + top border, thousands number format, % column as percent, outer border.
- ⬜ Pass ⬜ Fail

### T-W1.2 Undo formatting (FR-F5)
- Steps: after T-W1.1, click **Undo Formatting**.
- Expected: the table returns to its pre-format styling cell-by-cell.
- ⬜ Pass ⬜ Fail

### T-W1.3 Toggles (FR-F4)
- Steps: re-run Smart Format → **Toggles ▸ Result formats** (off), then again (on);
  repeat for **Number formats**.
- Expected: only that dimension flips each time; others unaffected.
- ⬜ Pass ⬜ Fail

### T-W1.4 Clear formatting
- Steps: select the table → **Clear Formatting**.
- Expected: all formatting removed; values intact.
- ⬜ Pass ⬜ Fail

### T-W1.5 Waterfall (FR-C1) — tab "Waterfall charts"
- Steps: select the EBITDA bridge (labels + values, 45→32→53) → **Charts ▸ Waterfall**.
- Expected: a stacked-column waterfall with an invisible base; increases green,
  decreases red, totals grey; data labels; no zero bars; "Base" hidden from legend.
- ⬜ Pass ⬜ Fail

### T-W1.6 Stacked Waterfall (FR-C2) — tab "Stacked Waterfall"
- Steps: select the FR/UK/DE multi-series bridge → **Stacked Waterfall**.
- Expected: anchors render full-height stacks; deltas float; categories coloured; legend shown.
- ⬜ Pass ⬜ Fail

### T-W1.7 CAGR arrow (FR-C4) — tab "CAGR Arrow"
- Steps: select the net-sales chart → **Display CAGR**.
- Expected: an arrow across the plot + a "CAGR +31.2%" label (9→15.5 over 2 years).
- ⬜ Pass ⬜ Fail

### T-W1.8 Theme-driven (FR-F6)
- Steps: copy `assets\theme.json`, change `colors.headerFill` to `#1F4E79`, set it as the
  theme in **Settings**; re-run Smart Format.
- Expected: header fill uses the new colour — no rebuild needed.
- ⬜ Pass ⬜ Fail

---

## Phase W2 — Modelling productivity

### T-W2.1 Autocolor (FR-M1)
- Steps: on a model range with inputs + formulas + a cross-sheet link, **Modelling ▸ Autocolor**.
- Expected: hardcoded inputs blue, in-sheet formulas black, cross-sheet/external refs green.
- ⬜ Pass ⬜ Fail

### T-W2.2 IFERROR (FR-M2)
- Steps: select formula cells → **IFERROR**.
- Expected: each formula becomes `=IFERROR(<orig>,"")`; already-wrapped/constants unchanged.
- ⬜ Pass ⬜ Fail

### T-W2.3 Fast Fill (FR-M3)
- Steps: put a formula in the top-left of a selection spanning right → **Fast Fill Right**;
  repeat down.
- Expected: relative refs shift per column/row; `$`-anchored refs stay; results match a manual fill.
- ⬜ Pass ⬜ Fail

---

## Phase W3 — Excel → PowerPoint live links

### T-W3.1 Export picture link (FR-X1)
- Steps: tab "Export Data to PowerPoint", select the income statement → **Export ▸ Export to PowerPoint**.
- Expected: a picture appears on a slide; PowerPoint comes to front.
- ⬜ Pass ⬜ Fail

### T-W3.2 Export native table (FR-X2)
- Steps: select a range → **Export as Table**.
- Expected: an editable PowerPoint table with the cell values.
- ⬜ Pass ⬜ Fail

### T-W3.3 Link Manager (FR-X4)
- Steps: PowerPoint → **Upslide ▸ Link Manager**.
- Expected: each link listed with slide #, type, source file, sheet!range, last refresh, status OK.
- ⬜ Pass ⬜ Fail

### T-W3.4 Refresh (FR-X5)
- Steps: change a number in Excel → in PowerPoint click **Refresh All** (or select + Refresh Selected).
- Expected: the slide object updates in place; position/size preserved; last-refresh updates.
- ⬜ Pass ⬜ Fail

### T-W3.5 Refresh from closed source
- Steps: close the Excel workbook → in PowerPoint **Refresh All**.
- Expected: the link still refreshes (source opened hidden), then status returns to OK.
- ⬜ Pass ⬜ Fail

### T-W3.6 Change source (FR-X6)
- Steps: rename/move the workbook → Link Manager → select link → **Change Source…** → pick the new file → Refresh.
- Expected: the link resolves to the new file; status OK.
- ⬜ Pass ⬜ Fail

---

## Phase W4 — Advanced linking & Word

### T-W4.1 Advanced Export (FR-X8) — tab "Advanced Export"
- Steps: build a small mapping table with headers `Range | Slide | Type` (e.g. two rows
  pointing at WACC ranges, slides 2 and 3) → select it → **Export ▸ Advanced Export**.
- Expected: each entry is exported as a linked object to its target slide in one action; re-runnable.
- ⬜ Pass ⬜ Fail

### T-W4.2 Sizing Guide (FR-X9)
- Steps: PowerPoint → **Sizing Guide ▸ Insert Placeholder** (note the key, e.g. `P1`); in the
  Advanced Export map add a `Placeholder` column = `P1` for one row → re-run.
- Expected: the exported object snaps to the placeholder's exact position/size.
- ⬜ Pass ⬜ Fail

### T-W4.3 Versioning / drift (FR-X7)
- Steps: change a linked source cell in Excel (do **not** refresh) → PowerPoint Link Manager →
  **Check Sources**.
- Expected: that link's status shows **Changed** (red); unchanged links show OK.
- ⬜ Pass ⬜ Fail

### T-W4.4 Highlight linked (FR-X10)
- Steps: in Excel (after exporting) → **Export ▸ Highlight Linked**.
- Expected: every range that has been linked out is filled/boxed.
- ⬜ Pass ⬜ Fail

### T-W4.5 Export to Word (FR-X3)
- Steps: open Word (blank doc) → in Excel select a range → **Export ▸ Export to Word**.
- Expected: a linked picture is inserted at the Word cursor.
- ⬜ Pass ⬜ Fail

### T-W4.6 Word Link Manager + Refresh (FR-X3)
- Steps: Word → **Upslide ▸ Link Manager** (link listed) → change the Excel cell → **Refresh All**.
- Expected: the Word object updates from source; status OK.
- ⬜ Pass ⬜ Fail

---

## Phase W5 — Library, settings, housekeeping

### T-W5.1 Save to Library (FR-L1)
- Steps: select a styled table → **Library ▸ Save to Library** → name it "IS".
- Expected: confirmation "Saved 'IS' …"; persists under `%APPDATA%\UpslideClone\library.json`.
- ⬜ Pass ⬜ Fail

### T-W5.2 Insert from Library (FR-L1)
- Steps: new sheet, select a cell → **Insert from Library** → pick "IS".
- Expected: the saved values are written starting at the selection.
- ⬜ Pass ⬜ Fail

### T-W5.3 Settings (FR-S1/S2)
- Steps: **Settings ▸** set a theme path + edit a shortcut value → Save → reopen Settings.
- Expected: values persisted across the dialog/Excel restart (`%APPDATA%\UpslideClone\settings.json`).
- ⬜ Pass ⬜ Fail

### T-W5.4 Clean & Prepare (FR-M6)
- Steps: add a defined name that points at a deleted cell (`#REF!`) → **Modelling ▸ Clean & Prepare**.
- Expected: the broken name is removed and reported; valid names remain.
- ⬜ Pass ⬜ Fail

### T-W5.5 Smart Print (FR-M7)
- Steps: **Modelling ▸ Smart Print**.
- Expected: print preview opens in landscape, fit-to-one-page-wide, narrow margins, centred.
- ⬜ Pass ⬜ Fail

### T-W5.6 Installer (W5)
- Steps: run `installer\uninstall.ps1`, reopen Excel (tab gone), run `installer\install.ps1`, reopen (tab back).
- Expected: clean unregister/register round-trip.
- ⬜ Pass ⬜ Fail

---

## PowerPoint design suite (open the training deck or any deck)

### T-D.1 Smart Align (8)
- Steps: select 3+ shapes → **Upslide ▸ Align & Size ▸ Smart Align ▸ Align Left** (try Center/Right/Top/Middle/Bottom).
- Expected: shapes align to the selection's bounding box; sizes unchanged. ⬜ Pass ⬜ Fail

### T-D.2 Distribute & Same Size (9)
- Steps: select 3+ shapes → **Distribute Across** / **Distribute Down**; then **Same Size**.
- Expected: equal gaps (ends fixed); all shapes resize to the first. ⬜ Pass ⬜ Fail

### T-D.3 Arrange (5)
- Steps: select a shape → **Arrange ▸ Bring to Front / Send to Back**; select 2+ → **Group / Ungroup**.
- Expected: z-order changes; grouping works. ⬜ Pass ⬜ Fail

### T-D.4 Format Shapes (6)
- Steps: select shapes/tables → **Format & Select ▸ Format Shapes**.
- Expected: brand font applied; table header gets the green fill + white bold. ⬜ Pass ⬜ Fail

### T-D.5 Select Similar (7) & Smart Painter (4)
- Steps: select one shape → **Select Similar** (selects same type+fill on the slide). Then select a styled shape + others → **Smart Painter**.
- Expected: similar shapes selected; first shape's format copied to the rest. ⬜ Pass ⬜ Fail

### T-D.6 Table of Contents (1) & Slide Check (10)
- Steps: **Slides ▸ Table of Contents** (inserts a TOC slide from titles); **Slide Check** (scan).
- Expected: TOC slide created; Slide Check reports off-slide shapes / missing titles. ⬜ Pass ⬜ Fail

### T-D.7 Outline pane (2)
- Steps: **Slides ▸ Outline** → double-click a title.
- Expected: pane lists slide titles; double-click jumps to the slide. ⬜ Pass ⬜ Fail

### T-D.8 Content Library (3)
- Steps: select shapes → **Content Library ▸ Save to Library** (name it). New slide → **Insert from Library** → pick it.
- Expected: shapes saved (to `Documents\UpslideClone`) and re-inserted. ⬜ Pass ⬜ Fail

### T-D.9 References (2)
- Steps: **References ▸ Footnote** (type text); **Cross-reference** (pick a slide). Re-order slides → **Refresh Refs**.
- Expected: numbered footnote added; cross-reference text shows the target's number/title and updates after re-order. ⬜ Pass ⬜ Fail

---

## Word

### T-WD.1 Format Table
- Steps: in a Word doc with a table, click inside it → **Upslide ▸ Formatting ▸ Format Table**.
- Expected: green header row (white bold), brand font, single borders. ⬜ Pass ⬜ Fail

### T-WD.2 Export → Word + Refresh (FR-X3)
- Steps: in Excel select a range → **Export to Word**; in Word **Upslide ▸ Link Manager**; change the Excel cell → **Refresh All**.
- Expected: linked object inserted, listed, and refreshed in place. ⬜ Pass ⬜ Fail

---

## Sign-off
| Phase | Cases | Pass | Fail | Notes |
|---|---|---|---|---|
| W1 | 8 | | | |
| W2 | 3 | | | |
| W3 | 6 | | | |
| W4 | 6 | | | |
| W5 | 6 | | | |
| PPT design | 9 | | | |
| Word | 2 | | | |

**Overall acceptance:** ⬜ Accept ⬜ Accept w/ fixes ⬜ Reject — _tester / date: ___________
