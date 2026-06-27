# UpslideClone — Technology Stack & Architecture (技术栈)

> How the whole add-in suite is built — languages, frameworks, project structure, tooling,
> and the key design decisions: the UI-free Core engine vs. thin VSTO shells, the
> Excel↔PowerPoint↔Word link engine, and the Office interop gotchas.

## 1. Platform & languages

| Layer | Choice | Why |
|---|---|---|
| Add-in model | **VSTO** (Visual Studio Tools for Office), COM-based | Only model with full cross-app Excel↔PowerPoint↔Word automation (Office.js can't do the linking). Windows-only — acceptable, Windows is primary. |
| Language | **C# 7.3** | VSTO's native language; strong typing for the link model + geometry. |
| Runtime | **.NET Framework 4.8.1** | Required by VSTO; matches installed targeting pack & Office. |
| Office automation | **Microsoft Office Interop** (Excel/PowerPoint/Word 15.0 PIAs), embedded interop types | Direct COM object-model access for ranges, charts, shapes, tables. |
| UI | **Ribbon XML** (IRibbonExtensibility) + **WinForms** task panes/dialogs | Full ribbon control; WinForms hosted via VSTO `CustomTaskPane`. |
| Tests | **xUnit** (SDK-style net481 project) | Fast, runs the UI-free core with no Office. |
| Serialization | **DataContractJsonSerializer** + **System.Xml.Linq** | In-box (no NuGet) for theme.json, settings, library, link registries. |

## 2. Solution structure

```
UpslideClone.sln
├─ src/UpslideClone.Core         ← UI-FREE logic (no Office reference)
│   ├─ Charts/      WaterfallEngine, StackedWaterfallEngine, CagrEngine
│   ├─ Formatting/  SmartFormatRules (result-row regex EN+CN, % detection)
│   ├─ Modelling/   A1, FormulaReferences (shift/extract), AutocolorClassifier,
│   │               FormulaTransform (IFERROR), DefinedNameAudit
│   ├─ Linking/     LinkMetadata(+tags), LinkHash, AdvancedExportMap,
│   │               LinkedRangeRegistry, LinkMetadataRegistry
│   ├─ Branding/    BrandTheme (theme.json loader)
│   ├─ Design/      AlignEngine, SlideCheckRules, TableOfContents, CrossReference
│   ├─ Library/     SnippetLibrary
│   ├─ Settings/    Settings, SettingsStore, ShortcutMap
│   └─ Util/        ColorUtil (BGR/OLE), Logger
├─ src/UpslideClone.Excel        ← Excel add-in (format, charts, modelling, export/linking, library, settings)
├─ src/UpslideClone.PowerPoint   ← PowerPoint add-in (Link Manager + refresh + 10-feature design suite)
├─ src/UpslideClone.Word         ← Word add-in (Link Manager + refresh + Format Table / Brand Font)
├─ tests/UpslideClone.Core.Tests ← xUnit (99 tests)
├─ assets/theme.json             ← editable branding
└─ installer/                    ← per-user install/uninstall PowerShell
```

**Core principle — the "engine vs dashboard" split (NFR-6):** all deterministic
logic lives in `Core` with **zero Office dependency**, so it is unit-tested on a
plain build server without Excel. The three add-ins are thin shells that read
ranges/shapes and call into Core. This is why 99 tests can prove the maths before
any Office app is opened, and why the Mac TS prototype and the Windows C# port
share one spec (NFR-9 parity).

## 3. How the linking engine works (the core value)

1. **Export (Excel):** copy the range as an EMF picture (`Range.CopyPicture`) or
   build a native table, drop it on a slide/doc, and tag it with link metadata.
2. **Metadata storage:**
   - *PowerPoint* — native `Shape.Tags` (`UPS_LinkId`, source workbook/sheet/range,
     type, hash, last-refresh).
   - *Word* — Word shapes have **no** Tags collection, so links live in a document
     **CustomXMLPart registry** (`LinkMetadataRegistry`) and each object is anchored
     by a bookmark `UPS_<id>`.
   - *Excel* — a `CustomXMLPart` registry of linked-out ranges drives "Highlight linked".
3. **Refresh:** enumerate tagged objects → resolve the source (open it **hidden** in a
   single cached Excel instance if closed) → re-render preserving geometry → update
   hash + timestamp.
4. **Drift (versioning):** a stable **SHA-256 content hash** of the source values is
   stored at export; the Link Manager recomputes it on demand and flags **Changed**.

## 4. Branding & theming

`assets/theme.json` (fonts 华文细黑/Calibri, brand palette, number formats,
result-row keywords) is loaded at runtime — edit without recompiling (FR-F6).
**Interop colour gotcha:** Office expects **BGR** integers, so every colour goes
through `ColorUtil.ToOle` (`ColorTranslator.ToOle`) — never raw RGB.

## 5. Reliability, logging, undo

- Every ribbon callback is wrapped: log → run → friendly message on failure; no
  unhandled exception reaches Office (NFR-3).
- Rolling log at `%APPDATA%\UpslideClone\logs` (NFR-8).
- Smart Format snapshots prior cell style to an in-memory **undo cache** (FR-F5).
- Strict COM lifetime management on the refresh path (release RCWs, close hidden Excel).

## 6. Build, sign, deploy

- **Build:** MSBuild from VS 2022's Office workload —
  `"D:\VS2022\MSBuild\Current\Bin\MSBuild.exe" UpslideClone.sln /t:Build /restore`.
  The VSTO project shells were generated from VS's own `VSTOExcel/PowerPoint/Word
  15 AddIn` templates, so the `ThisAddIn.Designer.cs` plumbing is version-exact.
- **Sign:** ClickOnce manifests signed with a per-user self-signed code cert
  (`CN=UpslideClone Dev`, CurrentUser\My); production should swap in a real cert.
- **Deploy:** `installer/install.ps1` trusts the cert (CurrentUser TrustedPublisher
  + Root) and registers each add-in under `HKCU\…\Office\<App>\Addins\…` with a
  `|vstolocal` manifest — a no-admin, per-user install. `uninstall.ps1` reverses it.

## 7. Toolchain summary

| Tool | Version / note |
|---|---|
| Visual Studio 2022 Community | `D:\VS2022`, Office/SharePoint + .NET desktop workloads |
| .NET Framework Dev Pack | 4.8.1 |
| .NET SDK | for the xUnit test project |
| MSBuild | 17.x (VS) — builds the VSTO projects (Build Tools alone can't) |
| Git + GitHub CLI | private repo `della119/UpslideClone` |
| Claude Code (Opus) | AI-assisted build, compiler-in-the-loop |

## 8. Why these choices (one line each)

- **VSTO over Office.js** — only path to Upslide-grade cross-app linking.
- **Core/shell split** — deterministic, testable, regression-safe (99 tests).
- **In-box serializers** — no NuGet restore friction on locked-down machines.
- **theme.json** — branding changes without redeploying code.
- **PowerShell installer** — internal, per-user, no admin; clean upgrade path to WiX.

## 9. Interop gotchas (hard-won — read before extending)

- **Dynamic narrowing.** With `EmbedInteropTypes=true`, COM members accessed off an
  indexer or an `object`-returning member are **late-bound (dynamic)**, and the runtime
  binder rejects `double→float` (`shape.Top = cell.Top`) and mis-casts COM SAFEARRAYs
  (`Object[*]→Object[]`, e.g. `series.Item(1).Values`). **Always cast off the indexer/Item
  first** (`((Excel.Range)sheet.Cells[r,c]).Top`, `(Excel.Series)sc.Item(1)`), then `(float)`.
- **Office colours are BGR.** Go through `ColorUtil.ToOle` / `OleFromHex` for any Excel/
  PowerPoint/Word colour property (`.Color`, `.RGB`, Word `WdColor`) — never raw RGB.
- **Word has no `Shape.Tags`.** Word links live in a document `CustomXMLPart` registry,
  anchored by bookmarks `UPS_<id>`; cross-references key on the stable `SlideID`.
- **PowerPoint Paste needs a window.** Pasting into a `WithWindow=false` presentation
  yields 0 shapes — the Content Library opens its backing deck **with** a window.
- **PowerPoint won't `SaveAs` into `%APPDATA%\Roaming`** (path-not-found) — the Content
  Library lives under `Documents\UpslideClone\`.
- **Namespace/alias collisions.** In the Word project, alias Word interop as `Wd` (not
  `Word`, which clashes with the `UpslideClone.Word` namespace); likewise `Settings` the
  class vs the namespace — alias in consumers.
- **Deploy hygiene.** After **any** rebuild, re-run `installer\install.ps1` (the signed
  manifests change, so the registration must be re-synced), then **cold-start** the app
  (`tools\Launch-*.cmd`) — Office's fast-restart keeps a background process that skips
  add-in init.
