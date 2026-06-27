# UpslideClone — Test Results (W1–W5 + PowerPoint design suite + Word)

_Generated: 2026-06-17 · latest update: 2026-06-24 · Configuration: Debug · .NET Framework 4.8.1 · Office x64_

## 1. Build matrix

| Project | Type | Result |
|---------|------|--------|
| UpslideClone.Core | .NET 4.8.1 class library (UI-free) | ✅ Build succeeded |
| UpslideClone.Core.Tests | xUnit (net481) | ✅ Build succeeded |
| UpslideClone.Excel | VSTO add-in | ✅ Build succeeded (.dll + .vsto + signed manifest) |
| UpslideClone.PowerPoint | VSTO add-in | ✅ Build succeeded (.dll + .vsto + signed manifest) |
| UpslideClone.Word | VSTO add-in | ✅ Build succeeded (.dll + .vsto + signed manifest) |

Full-solution build (`msbuild UpslideClone.sln /t:Build /restore`): **0 errors.**

## 2. Automated unit tests — `UpslideClone.Core.Tests`

**Result: ✅ 99 passed / 0 failed / 0 skipped.**

These exercise the pure, deterministic core logic (no Office required), which is the
parity oracle shared with the Mac Office.js prototype.

| Area (phase) | Representative tests | What's verified |
|---|---|---|
| Waterfall geometry (W1) | first-row anchor, increase floats on running total, decrease lowers base, anchor auto-detect on 45→32→53 | Bridge bar math matches the training fixture |
| Stacked waterfall (W1) | anchor full-height, positive step float, negative step floor drop | Per-category stacking geometry |
| CAGR (W1) | 9→15.5 over 2 periods = 31.2%, negative growth label, guard on non-positive base | Compound-growth math + label |
| Smart Format rules (W1) | result-row regex EN+CN (EBITDA/Gross Margin/毛利/净利润/营业利润/合计), % column detection | Auto result-row + percent detection |
| Brand theme (W1) | brand defaults, theme.json parse, hex→BGR/OLE | Theme loading + Interop colour conversion |
| Autocolor (W2) | constant=blue, formula=black, cross-sheet/external=green | Banker colour classification |
| IFERROR (W2) | wrap body, custom replacement, no double-wrap, leaves constants | Formula transform |
| A1 / Fast Fill (W2) | column round-trips (A↔1, XFD↔16384), relative shift, $-anchor stays, strings/sheets/funcs untouched | Reference shifter (fill / paste-preserve) |
| Smart Track refs (W2) | extract A1 refs, ignore string literals | Precedent reference extraction |
| Link model (W3) | LinkHash deterministic + change-sensitive, null≠"", Shape.Tags round-trip | Live-link metadata + drift hash |
| Advanced Export (W4) | header-keyed map parse, defaults, missing-range error | Range→slide mapping table |
| Highlight registry (W4) | upsert by id, XML round-trip | Workbook linked-range registry |
| Word doc registry (W4) | upsert/replace, XML round-trip, bookmark naming | Word link store (no Shape.Tags) |
| Settings (W5) | theme+overrides round-trip, effective-shortcut overlay, garbage→defaults | Per-user settings persistence |
| Excel Library (W5) | upsert by name (case-insensitive), JSON round-trip, remove | Snippet library model |
| Clean & Prepare (W5) | #REF! detection, external-link detection | Defined-name audit |
| Align/Distribute (PPT design) | align Left/Right/Center to bbox, equal-gap distribute, match-size, ends-fixed | Shape-layout geometry |
| Slide Check (PPT design) | off-slide detection, tiny-text floor | Pre-send QC rules |
| Table of Contents (PPT design) | skip untitled/skipped, numbered render | TOC builder |
| Cross-reference (PPT design) | format with/without title, footnote numbering | Reference text helpers |

## 2b. Headless live-run on the training files (2026-06-22)

Driven via `tools\Run-Demo.ps1` (same tested Core logic + theme the add-in uses)
against `Training Guide Excel_test.xlsx` / `..._test.pptx`, then verified
visually by screenshot:

| Operation | Result | Visual check |
|---|---|---|
| Smart Format (Format tables) | green header, 1-decimal numbers, (5.7%) parens, EBITDA/Gross Margin emphasised | ✅ matches "Expected Result" (PPT slide) |
| Waterfall (Waterfall charts) | dark-green anchors (45/32), red/green deltas, (10) labels | ✅ matches Expected |
| CAGR (CAGR Arrow) | CAGR +31.2% over 2 periods | ✅ computed + arrow drawn |
| Stacked Waterfall | 6 steps × FR/UK/DE, floating stacked deltas | ✅ renders (mixed-sign geometry simplified per W4 note) |
| Autocolor (Advanced Export WACC) | 26 inputs blue, 52 formulas black | ✅ applied |
| IFERROR (Advanced Export) | 17 formulas wrapped | ✅ applied |
| Export → PowerPoint | linked picture + UPS_LinkId tags, new slide | ✅ slide shows formatted IS |

**Theme correction made during this run:** header `#53565A`→`#A9D18E` (green),
number format → 1 decimal, percent → parens on negatives, waterfall total/anchor
`#53565A`→`#375623` (dark green) — to match the firm's Excel standard + the
training "Expected Result". Applied in `BrandTheme.Default()` and `assets/theme.json`.

## 2c. Headless live-test — PowerPoint design suite & Word (2026-06-22)

Driven via `tools\Run-PPT-Design-Demo.ps1` (design ops on real shapes) + targeted
COM smokes, using the same Core engines the ribbon buttons call:

| Operation | Result |
|---|---|
| Smart Align (Left) on 4 shapes | ✅ PASS — all share min Left |
| Distribute (Down) | ✅ PASS — equal gaps, ends fixed |
| Same Size | ✅ PASS — match first shape |
| Table of Contents | ✅ PASS — 2 titled slides → TOC |
| Slide Check | ✅ PASS — off-slide shape detected |
| **Content Library** save → insert round-trip | ✅ PASS (after fixing: paste needs a windowed presentation; library moved to `Documents\UpslideClone` since PowerPoint won't SaveAs into `%APPDATA%`) |
| **Word** Format Table header shading | ✅ PASS — WdColor set/read = #A9D18E |

## 3. Integration / cross-app status

The Office Interop layers (Excel/PowerPoint/Word automation) are **compile-verified**
and the add-ins are **built, signed, registered, and loadable**. Their runtime
behaviour is exercised by the **manual UAT plan** (`UAT_PLAN.md`) — not by automated
tests, because driving live Office across apps is outside the unit-test scope.

| Capability | Code status | Runtime status |
|---|---|---|
| Smart Format / toggles / undo (W1) | ✅ built | ⏳ UAT |
| Waterfall / Stacked / CAGR (W1) | ✅ built | ⏳ UAT |
| Autocolor / IFERROR / Fast Fill (W2) | ✅ built | ⏳ UAT |
| Export → PPT (picture/table) + Link Manager + Refresh (W3) | ✅ built | ⏳ UAT |
| Advanced Export / Sizing Guide / Drift / Highlight (W4) | ✅ built | ⏳ UAT |
| Export → Word + Word Link Manager / Refresh (W4) | ✅ built | ⏳ UAT |
| Library / Settings / Clean & Prepare / Smart Print (W5) | ✅ built | ⏳ UAT |
| PPT design suite — Align/Distribute/Arrange/Format/Select/Painter/TOC/SlideCheck | ✅ built | ✅ headless 5/5 |
| PPT Content Library + Cross-refs/Footnotes/Outline | ✅ built | ✅ library round-trip; refs ⏳ UAT |
| Word formatting (Format Table / Brand Font) | ✅ built | ✅ shading smoke |
| Installer (per-user register + trust) | ✅ verified running | ✅ ran clean |

## 4. Known limitations (to harden post-UAT)

- **Keyboard shortcuts (FR-S2):** configurable + persisted in Settings, but live
  key-binding via `Application.OnKey` is **not yet wired** (needs a managed key
  bridge / hook). Ribbon screentips show the intended chords.
- **Picture refresh** re-pastes via the clipboard (brief clipboard use).
- **Table refresh / export** copies cell **display text**, not number formats.
- **Drift status** is on-demand (Link Manager → Check Sources), not automatic.
- **Installer** is per-user dev-cert based; production needs a real code-signing
  cert + ClickOnce/WiX (see `installer/README.md`).

## 5. How to reproduce
```powershell
# Unit tests
dotnet test tests\UpslideClone.Core.Tests\UpslideClone.Core.Tests.csproj -c Debug

# Full build
& "D:\VS2022\MSBuild\Current\Bin\MSBuild.exe" UpslideClone.sln /t:Build /restore
```
