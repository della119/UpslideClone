# UpslideClone — Build Status

> Phase-by-phase build log — from the W0 scaffold through the W3 linking engine, the
> PowerPoint design suite, and the Word add-in, to the **v1.0 pre-built release**.

_Last updated: 2026-06-27._

## ✅ Verified on this machine (light tier)
- **.NET 4.8.1 targeting pack + .NET SDK 8.0.422** installed.
- `UpslideClone.Core` **builds clean** (0 warnings / 0 errors).
- `UpslideClone.Core.Tests` — **55/55 xUnit tests pass** (W1 + W2 logic).
- Projects retargeted **v4.8 → v4.8.1** (the Dev Pack winget installed was 4.8.1; it's a binary-compatible superset and runs under Office — minor deviation from NFR-1's "4.8", revisit if a 4.8 pack is later required).


## What's done (authored, in repo)

### `UpslideClone.Core` (pure .NET 4.8 — no Office) — **W1 logic ported from the Mac TS prototype**
- `Charts/WaterfallEngine.cs` — faithful port of `addin/src/core/waterfall.ts` (anchor detection, base float, inc/dec/total). Pure + unit-tested.
- `Charts/StackedWaterfallEngine.cs` — FR-C2 multi-series geometry (shared Base float + per-category segments).
- `Charts/CagrEngine.cs` — FR-C4 CAGR + arrow label.
- `Formatting/SmartFormatRules.cs` — port of `smartFormat.ts` (`RESULT_RE` EN+CN, `isPercentColumn`).
- `Branding/BrandTheme.cs` — theme.json loader (uses `DataContractJsonSerializer`, **no NuGet**).
- `Settings/ShortcutMap.cs`, `Settings/Settings.cs` — Appendix C shortcut map, per-user settings model.
- `Linking/LinkMetadata.cs` — W3 link schema + `UPS_*` tag keys.
- `Util/ColorUtil.cs` (BGR/OLE), `Util/Logger.cs` (rolling log).

### `UpslideClone.Core.Tests` (xUnit) — geometry, % detection, result-row regex (EN+CN), CAGR, theme, color.

### `UpslideClone.Excel` (VSTO add-in source — **W1 commands**)
- `Ribbon/UpslideRibbon.xml` + `UpslideRibbon.cs` — Upslide tab (Formatting / Charts / Modelling / Export / Settings groups), wrapped callbacks (NFR-3).
- `Commands/SmartFormatCommand.cs` — Smart Format + toggles (title/result/number/percent) + Clear.
- `Commands/UndoCache.cs` — FR-F5 reversible undo (in-memory cell-style snapshots).
- `Commands/BuildWaterfallCommand.cs` — Appendix A.2 Interop port.
- `Commands/BuildStackedWaterfallCommand.cs`, `Commands/CagrCommand.cs`.
- `Commands/ExcelHelpers.cs`, `Commands/ThemeProvider.cs`, `ThisAddIn.cs`.

### `UpslideClone.Core` — W2 modelling logic (pure, **tested on this machine**)
- `Modelling/AutocolorClassifier.cs` — FR-M1 banker colour classes (input/formula/link).
- `Modelling/FormulaTransform.cs` — FR-M2 IFERROR wrapper (no double-wrap).
- `Modelling/FormulaReferences.cs` + `A1.cs` — relative/absolute A1 shifter (Fast Fill / paste-preserve) and reference extractor (Smart Track), string/sheet/function-safe.

### `UpslideClone.Excel` — W2 commands (drop-in, build needs VSTO)
- `Commands/AutocolorCommand.cs`, `IferrorCommand.cs`, `FastFillCommand.cs`; ribbon Modelling group wired.

### Assets / solution
- `assets/theme.json` (brand default), `UpslideClone.sln` (Core + Tests).

### Still pending VSTO/heavy-tier (later)
- Smart Track task pane (UI), Explorer pane, keyboard-shortcut registration (Application.OnKey), and all of W3+ (linking engine). W3 is the core moat.

## ✅ BUILD UNBLOCKED — full toolchain installed (2026-06-17)

VS 2022 Community @ `D:\VS2022` (Office/SharePoint + .NET desktop workloads), .NET 4.8.1 Dev Pack, .NET SDK. **The whole solution now compiles**, including the VSTO add-in:
- `UpslideClone.Core.dll` + **55/55 tests green**
- `UpslideClone.Excel.dll` + `UpslideClone.Excel.vsto` + signed `.dll.manifest` (sideload-ready)
- VSTO project shell authored from VS's own template (`ThisAddIn.Designer.cs`, blueprint xml, csproj) — all 3 projects target **v4.8.1**.
- ClickOnce manifests signed with a **dev-only self-signed cert** (CurrentUser\My, thumbprint `1EF116…9CC2`); PFX is gitignored. Regenerate per machine: `New-SelfSignedCertificate -Type CodeSigningCert -Subject "CN=UpslideClone Dev" -CertStoreLocation Cert:\CurrentUser\My -KeyExportPolicy Exportable` then update `ManifestCertificateThumbprint` in the csproj.

### W3 linking engine — built & compiling (2026-06-17)
Excel→PowerPoint live links, all 4 projects compile, 62/62 tests:
- **Core:** `Linking/LinkHash.cs` (source-drift hash), `LinkMetadata.ToTags/FromTags` (Shape.Tags round-trip). Tested.
- **Excel:** `Commands/ExportToPowerPointCommand.cs` — Export selection to PPT as **linked picture** (FR-X1) or **native table** (FR-X2), tagged with UPS_* metadata. Ribbon: Export to PowerPoint / Export as Table.
- **PowerPoint:** `UpslideClone.PowerPoint` VSTO add-in — `RefreshEngine` (FR-X5, re-render preserving geometry), `LinkManagerService` (list / change-source FR-X6 / go-to), `Panes/LinkManagerControl` (FR-X4 task pane), Links ribbon group.

### To run it — TWO add-ins now (Excel + PowerPoint)
For the full export→refresh loop, both add-ins must be loaded. In VS: set **Solution** startup to run both, or F5 each. Simplest: right-click solution → Properties → Startup Project → **Multiple startup projects** → set `UpslideClone.Excel` and `UpslideClone.PowerPoint` to **Start** → **F5**.

End-to-end test (W3 exit criterion):
1. In Excel (`Training Guide Excel.xlsx`), select the **Export Data to PowerPoint** tab's Income Statement range → **Upslide ▸ Export to PowerPoint** (or Export as Table). A linked object lands on a slide.
2. In PowerPoint, **Upslide ▸ Link Manager** → see the link listed (source, sheet!range, status).
3. Change a number in Excel → in PowerPoint click **Refresh All** → the slide object updates in place.
4. Move/rename the source → **Change Source…** → repoint → Refresh.

> W3 picture refresh re-pastes (brief clipboard use); table refresh copies cell *text* (display values), not formats.

### W4 advanced linking — W4a built & compiling (2026-06-17)
All 4 projects compile, 68/68 tests:
- **Advanced Export (FR-X8):** `Core/AdvancedExportMap` (header-keyed map parser, tested) + Excel `AdvancedExportCommand` (batch many ranges→slides from a selected mapping table). Ribbon: Advanced Export.
- **Sizing Guide (FR-X9):** PowerPoint `SizingGuideCommand` inserts tagged placeholder rectangles; export snaps objects to a named placeholder (`ExportRangeToSlide` + Advanced Export's Placeholder column). PPT ribbon: Insert Placeholder.
- **Versioning / drift (FR-X7):** `LinkManagerService.List(pres, checkDrift)` opens sources and compares live hash to stored → **OK / Changed / Source missing**. Link Manager pane: **Check Sources** button.
- **Highlight linked (FR-X10):** `Core/LinkedRangeRegistry` (CustomXMLPart, tested) + Excel `LinkRegistryStore` (records on export) + `HighlightLinkedCommand` (flags linked-out ranges). Ribbon: Highlight Linked.

### W4b — Word export (FR-X3) — built & compiling (2026-06-17)
All **5 projects** compile, 72/72 tests:
- **Core:** `LinkMetadataRegistry` (document-level link store as XML, tested) — Word has no `Shape.Tags`, so links live in a `CustomXMLPart` and are anchored by bookmarks `UPS_<id>`.
- **Excel:** `ExportToWordCommand` (FR-X3) inserts a linked picture/table at Word's cursor, bookmarks it, writes the doc registry. Ribbon: Export to Word.
- **Word add-in (`UpslideClone.Word`):** `RefreshEngine` (re-paste picture / rewrite table in the bookmark range), `LinkManagerService` (list / change-source / go-to), `LinkManagerControl` pane, Links ribbon. Signed with the same dev cert; registered in HKCU.

**W4 COMPLETE (code).** All registered for testing (Excel + PowerPoint + Word).

### W5 — Library, Settings, housekeeping, installer — built & compiling (2026-06-17)
**All 5 projects compile, 85/85 tests.** Feature-complete vs the spec:
- **Excel Library (FR-L1):** `Core/SnippetLibrary` (tested) + Excel `ExcelLibraryCommand` (Save to / Insert from). Ribbon: Library group.
- **Settings (FR-S1/S2):** `Core/SettingsStore` (tested) + Excel `SettingsForm`/`SettingsCommand` (theme path + shortcut grid, persisted to `%APPDATA%`). Ribbon: Settings.
- **Clean & Prepare (FR-M6):** `Core/DefinedNameAudit` (tested) + Excel `CleanPrepareCommand` (removes #REF! names).
- **Smart Print (FR-M7):** Excel `SmartPrintCommand` (landscape/fit-to-width/preview).
- **Installer:** `installer/install.ps1` + `uninstall.ps1` + README — per-user register + cert trust, **verified running clean**.

**ALL PHASES W0–W5 COMPLETE IN CODE.** Deliverables: `TEST_RESULTS.md`, `UAT_PLAN.md`, `TECH_STACK.md`.

### Phase D — PowerPoint design suite + Word formatting (2026-06-22)
All 5 projects build, **99/99 tests**.
- **PowerPoint design suite (all 10 Upslide PPT features):** Smart Align, Resize & Distribute, Arrange, Format shapes/text/tables, Select Similar, Smart Painter, Table of Contents, Slide Check, **Content Library**, **Cross-references / Footnotes / Outline pane**. Core engines (`AlignEngine`, `SlideCheckRules`, `TableOfContents`, `CrossReference`) tested; 8 geometry/format ops headless-verified 5/5; Content Library round-trip PASS.
- **Word formatting toolkit:** `Format Table` (brand the selected table), `Brand Font`. Header-shading smoke PASS.
- **README.md + USER_GUIDE.md** added; all docs refreshed.

**Remaining:** full manual UAT (run `UAT_PLAN.md` — now incl. PPT design + Word), then optionally harden known limits (live OnKey shortcuts; production WiX/cert). **Project is feature-complete vs. the Upslide Excel + PowerPoint training guides.**

---

## (historical) Blocked — toolchain not set up (2026-06-16)

| Check | Result |
|-------|--------|
| Visual Studio IDE | ❌ Only **Build Tools 2022** present (no IDE → no F5 sideload/debug) |
| Office/SharePoint dev workload | ❌ No VSTO build targets, no project templates |
| .NET Framework 4.8 targeting pack | ❌ `MSB3644` — Core won't build until installed |
| .NET SDK (`dotnet`) | ❌ Not installed — needed for the SDK-style test project |
| VSTO runtime / Office | ✅ Present; Office is **x64**; Excel installed |

**Consequence:** nothing can be compiled or F5-debugged on this machine yet, and the two VSTO project _shells_ (`.csproj` + `ThisAddIn.Designer.cs` + manifests) are normally generated by the VS "Excel/PowerPoint VSTO Add-in" template — which requires the Office workload. The hand-written C# above drops into those shells unchanged.

## Next steps (pending your go-ahead on environment)
1. Install **Visual Studio 2022 Community** with **Office/SharePoint development** + **.NET desktop development** workloads (pulls in the .NET 4.8 targeting pack, VSTO templates/targets, and the .NET SDK).
2. In VS: create `UpslideClone.Excel` + `UpslideClone.PowerPoint` VSTO add-in projects; drop in the provided `Ribbon/`, `Commands/`, `ThisAddIn.cs`; add a Core project reference.
3. Build → set Excel project as startup → **F5** → test against `Training Guide Excel.xlsx` (Format tables / Waterfall / Stacked Waterfall / CAGR tabs).
