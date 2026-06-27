# Upslide Clone — Excel Add-in (Phase 1)

Cross-platform Office.js add-in reproducing Upslide's **Smart Format** and **Waterfall** features.
This is the **Mac deliverable** *and* the proven prototype that gets ported to C#/VSTO for the
full Windows product (Phase 2: PowerPoint linking engine).

## What's implemented
- **Smart Format (tables)** — branded header, result-row detection (EBITDA, Gross Margin…),
  number/percent formats, alignment, borders. Ribbon button + task pane.
- **Waterfall builder** — stacked-column waterfall with invisible "Base" float series and
  auto-detected anchor/total bars. `#N/A` trick suppresses zero bars/labels.

## Project layout
```
manifest.xml            Add-in manifest + "Upslide" ribbon tab
src/taskpane/           Task-pane UI (html/css/ts)
src/commands/           Ribbon button handlers
src/core/smartFormat.ts Smart Format engine (pure logic, ports to C#)
src/core/waterfall.ts   Waterfall geometry + chart build (ports to C#)
assets/                 Icons
```

## Run & test (Excel for Mac)
```bash
cd addin
npm install        # once
npm start          # builds, trusts dev cert, sideloads into Excel desktop, starts dev server
```
`npm start` opens Excel and registers the add-in. Look for the **Upslide** tab on the ribbon.

To test: open `../Training Guide Excel.xlsx`, go to the **Format tables** or
**Waterfall charts** tab, select the sample table, and click the matching ribbon button.

Stop the dev session:
```bash
npm stop
```

## Other commands
```bash
npm run build      # production bundle (dist/)
npm run dev-server # just the webpack https dev server on :3000
npm run validate   # validate manifest.xml
```

## Notes / next
- Colors use a brand palette (green increase / red decrease / gray total); will be
  made configurable in the task pane.
- Undo cache (Upslide-style reversible formatting) not yet implemented.
- Next Phase-1 features: CAGR arrow, Autocolor, IFERROR wrapper, Fast Fill.
- Phase 2 (Windows/VSTO): Excel→PPT/Word linking, Link Manager, Advanced Export.
