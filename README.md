# UpslideClone — In-house Office Productivity Add-in Suite

A Windows VSTO add-in suite that replaces the commercial **Upslide** licence: branded
Excel formatting, finance charts, an Excel↔PowerPoint↔Word **live-link engine**, and a
PowerPoint slide-design toolkit — built in-house, on the firm's brand template
standard.

> **New here?** Jump to [Quick start](#quick-start) · [User Guide](USER_GUIDE.md) ·
> [Tech stack](TECH_STACK.md) · [Test results](TEST_RESULTS.md) · [UAT plan](UAT_PLAN.md)

---

## 1. Business case

| | |
|---|---|
| **Problem** | Upslide is a paid, per-seat licence. Its value is the Excel→PowerPoint linking + branded formatting workflow used daily for client decks. |
| **Solution** | Re-implement that workflow as an in-house add-in suite — full ownership, no per-seat cost, tailored to the firm's brand template (华文细黑 / Calibri, green headers `#A9D18E`, RMB ¥). |
| **Owner** | Della (finance + tech) |
| **Outcome** | A user formats a table/chart, exports it as a **linked** object to PowerPoint/Word, changes the Excel data, and **refreshes** the deck — without Upslide installed. |

**Why VSTO and not Office.js:** only the COM object model (VSTO) supports the deep,
cross-application automation the live-link engine needs. Office.js can't drive
PowerPoint/Word from Excel. VSTO is Windows-only — acceptable, as Windows is the primary
environment. (A cross-platform formatting/chart subset exists as the Mac Office.js
prototype in `addin/`.)

---

## 2. What it does (feature inventory)

### Excel add-in (`UpslideClone.Excel`)
- **Smart Format** — one-click branded tables: green header, result-row emphasis
  (EBITDA / Gross Margin / 毛利 / 净利润 …), number/percent formats, borders; reversible
  **undo** + per-dimension **toggles**.
- **Charts** — **Waterfall**, **Stacked Waterfall**, **CAGR arrow** (floating-bar bridge
  reproduced from a stacked column chart, dark-green anchors).
- **Modelling** — **Autocolor** (blue inputs / black formulas / green links), **IFERROR**
  wrapper, **Fast Fill** (structure-aware), **Clean & Prepare** (#REF! name pruning),
  **Smart Print**.
- **Export & linking** — **Export to PowerPoint / Word** (linked picture or native table),
  **Advanced Export** (batch many ranges→slides), **Highlight Linked**.
- **Library / Settings** — reusable snippet library; theme + shortcut settings.

### PowerPoint add-in (`UpslideClone.PowerPoint`)
- **Links** — **Link Manager** pane (list / status / change-source / go-to), **Refresh
  All** (re-render from Excel, geometry preserved), **drift** detection (Check Sources).
- **Sizing Guide** — tagged placeholders so exports land at consistent dimensions.
- **Design toolkit** — Smart Align, Resize & Distribute, Arrange, Format shapes/text/tables,
  Select Similar, Smart Painter, Table of Contents, Slide Check, **Content Library**,
  **Cross-references / Footnotes / Outline pane**.

### Word add-in (`UpslideClone.Word`)
- **Links** — Link Manager + Refresh for Excel→Word linked objects (anchored by bookmarks
  + a document link registry, since Word shapes have no `Shape.Tags`).
- **Formatting** — **Format Table** (brand the selected table: green header, fonts, borders)
  and **Brand Font**.

---

## 3. Architecture (one idea to remember)

**Engine vs. dashboard.** All deterministic logic lives in **`UpslideClone.Core`** — a
UI-free .NET library with **zero Office dependency**. The three add-ins are thin shells
that read ranges/shapes and call into Core. This is why **85 unit tests** prove the maths
(waterfall geometry, % detection, CAGR, formula shifting, link hashing, alignment) on a
plain build server, with no Office open.

```
UpslideClone.sln
├─ src/UpslideClone.Core         ← UI-free logic + tests target (Charts, Formatting,
│                                   Modelling, Linking, Branding, Library, Settings, Design)
├─ src/UpslideClone.Excel        ← Excel VSTO add-in (ribbon + commands + export)
├─ src/UpslideClone.PowerPoint   ← PowerPoint VSTO add-in (Link Manager, refresh, design)
├─ src/UpslideClone.Word         ← Word VSTO add-in (Link Manager, refresh)
├─ tests/UpslideClone.Core.Tests ← xUnit (85+ tests)
├─ assets/theme.json             ← editable branding (fonts, colours, number formats)
├─ installer/                    ← per-user install / uninstall (PowerShell)
├─ tools/                        ← cold-start launchers, headless demo runner
└─ addin/                        ← Mac Office.js prototype (parity oracle)
```

The **link engine**: exported objects are tagged (`UPS_LinkId`, source workbook/sheet/
range, type, **SHA-256 content hash**, last-refresh). Refresh re-opens the source (hidden,
cached Excel instance if closed), re-renders preserving position/size, and re-hashes for
drift. See [TECH_STACK.md](TECH_STACK.md) §3.

---

## 4. Quick start

### Two ways to install (per-user, no admin)

**🟢 Non-developers — no Visual Studio needed.** Download the latest zip from
**[Releases](https://github.com/della119/UpslideClone/releases)**, extract it, and
double-click **`Install.cmd`**. It trusts the bundled certificate, installs the
already-compiled add-ins, and confirms they load. Requirements: Windows 10/11 64-bit +
desktop **Office 64-bit** (Excel/PowerPoint/Word) + the VSTO Runtime (normally already
installed with Office). To remove: `Uninstall.cmd`.

> **⚠️ Security note — self-signed certificate.** `Install.cmd` adds the bundled
> certificate (`UpslideClone.cer`, public key only) to your **Trusted Publishers** and
> **Root** stores so Office will load the add-ins without a prompt. This is unavoidable for
> self-signed VSTO add-ins — Windows won't trust them otherwise — and is fine for
> **internal / known-source** use. For wide distribution to people who don't know you,
> re-sign the add-ins with a **CA-issued code-signing certificate** instead, so no manual
> trust step is needed and the publisher is verifiable.

**🛠 Build from source.** Prerequisites: **Visual Studio 2022** with the *Office/SharePoint
development* + *.NET desktop* workloads (pulls in the .NET 4.8.x targeting pack + VSTO build
targets). Then just double-click **`installer\Install-Upslide.cmd`** — it auto-creates the
signing cert, builds Release, copies to a stable location, registers, and verifies each
add-in loads. Equivalent:
```powershell
powershell -ExecutionPolicy Bypass -File installer\Install-Upslide.ps1
```
Then **open Excel / PowerPoint / Word normally** — the **Upslide** ribbon tab loads. If it
ever disappears, double-click `installer\Repair-Upslide.cmd`. See
[installer/README.md](installer/README.md) for details and the dev (`bin\Debug`) workflow.

### Run the tests
```powershell
dotnet test tests\UpslideClone.Core.Tests\UpslideClone.Core.Tests.csproj -c Debug
```

### Don't open the `.sln` by double-click
It can launch in *Blend* (a XAML tool that can't load Office projects). Open **Visual
Studio 2022** first, then *File ▸ Open ▸ Solution*. You don't need an IDE to build or test.

---

## 5. How it was built

- **Language/runtime:** C# 7.3 on **.NET Framework 4.8.1** (VSTO requirement).
- **VSTO project shells** were generated from Visual Studio's own Excel/PowerPoint/Word
  add-in templates, so the version-sensitive `ThisAddIn.Designer.cs` plumbing is exact.
- **Compiler-in-the-loop:** every phase = author Core logic → unit-test → author the
  Interop commands + Ribbon XML → MSBuild → fix → register → test live in Office.
- **Phased delivery:** W0 scaffold → W1 formatting/charts → W2 modelling → W3 linking
  engine → W4 advanced linking + Word → W5 library/settings/installer → PowerPoint design
  suite. See [PROJECT_STATUS.md](PROJECT_STATUS.md).
- **Signing/deploy:** ClickOnce manifests signed with a per-user dev cert; `install.ps1`
  trusts it and registers under `HKCU`. Production should swap in a real code-signing cert
  + WiX/ClickOnce.

---

## 6. Repository docs

| File | What |
|---|---|
| [USER_GUIDE.md](USER_GUIDE.md) | End-user guide — every button, what it does, how to use it |
| [TECH_STACK.md](TECH_STACK.md) | Technology + architecture deep-dive |
| [PROJECT_STATUS.md](PROJECT_STATUS.md) | Phase-by-phase build status |
| [TEST_RESULTS.md](TEST_RESULTS.md) | Build matrix + unit/integration results |
| [UAT_PLAN.md](UAT_PLAN.md) | Step-by-step manual acceptance tests |
| [UPSLIDE_CLONE_Windows_Requirements.md](UPSLIDE_CLONE_Windows_Requirements.md) | Original requirements |
| [Upslide_Functionality_Summary.md](Upslide_Functionality_Summary.md) | Reverse-engineered feature reference |

---

## 7. Status & roadmap

- ✅ **W0–W5 complete** (Excel + PowerPoint + Word linking, formatting, charts, modelling,
  library, settings, installer) · live-verified on the training files (Smart Format,
  Waterfall, Stacked Waterfall, CAGR, Autocolor, IFERROR, Export→PPT).
- ✅ **PowerPoint design suite — all 10** (Smart Align, Resize & Distribute, Arrange, Format,
  Select Similar, Smart Painter, Table of Contents, Slide Check, Content Library,
  Cross-references/Footnotes/Outline) · 8 geometry/format ops headless-verified 5/5;
  Content Library round-trip PASS.
- ✅ **Word formatting toolkit** (Format Table, Brand Font) + Link Manager/Refresh.
- ✅ **99/99 unit tests green** · all 5 projects build · signed + registered · on GitHub.
- ✅ **Production per-user installer** (`installer\Install-Production.ps1`) — deploys to a
  stable location so the tab loads on a normal launch (no cold-start launcher needed).
- ⬜ **Remaining** — full manual click-through UAT (`UAT_PLAN.md`); enterprise distribution
  (real code-signing cert + WiX/MSI for machine-wide push); live keyboard-shortcut binding.

**Feature-complete vs the Upslide Excel + PowerPoint training guides.**

---

## 8. Licence / IP note

In-house tooling. The proprietary **Upslide training fixtures** are git-ignored and
never committed — they are not ours to redistribute. The code and reverse-engineered specs
are the firm's.
