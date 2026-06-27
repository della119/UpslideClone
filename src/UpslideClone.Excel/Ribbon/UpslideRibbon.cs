using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using UpslideClone.Core.Util;
using UpslideClone.Excel.Commands;

namespace UpslideClone.Excel.Ribbon
{
    /// <summary>
    /// Ribbon-XML extensibility object for the Upslide tab. Each onAction callback
    /// is wrapped so no exception ever reaches Office (NFR-3): failures are logged
    /// and surfaced as a non-blocking message.
    /// </summary>
    [System.Runtime.InteropServices.ComVisible(true)]
    public sealed class UpslideRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbon;

        public string GetCustomUI(string ribbonID)
        {
            return ReadResource("UpslideClone.Excel.Ribbon.UpslideRibbon.xml");
        }

        public void OnRibbonLoad(Office.IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
            Logger.Info("Upslide ribbon loaded.");
        }

        // ---- Formatting ----
        public void OnSmartFormat(Office.IRibbonControl c) => Run("Smart Format", () => SmartFormatCommand.Run());
        public void OnUndoFormatting(Office.IRibbonControl c) => Run("Undo Formatting", () => UndoCache.UndoLast());
        public void OnClearFormatting(Office.IRibbonControl c) => Run("Clear Formatting", () => SmartFormatCommand.ClearFormatting());

        public void OnToggleTitle(Office.IRibbonControl c) => Run("Toggle title", () => SmartFormatCommand.Toggle(SmartFormatDimension.Title));
        public void OnToggleResult(Office.IRibbonControl c) => Run("Toggle result", () => SmartFormatCommand.Toggle(SmartFormatDimension.Result));
        public void OnToggleNumber(Office.IRibbonControl c) => Run("Toggle number", () => SmartFormatCommand.Toggle(SmartFormatDimension.NumberFormat));
        public void OnTogglePercent(Office.IRibbonControl c) => Run("Toggle percent", () => SmartFormatCommand.Toggle(SmartFormatDimension.PercentFormat));

        // ---- Modelling (W2) ----
        public void OnAutocolor(Office.IRibbonControl c) => Run("Autocolor", () => AutocolorCommand.Run());
        public void OnIferror(Office.IRibbonControl c) => Run("Apply IFERROR", () => IferrorCommand.Run());
        public void OnFastFillRight(Office.IRibbonControl c) => Run("Fast Fill Right", () => FastFillCommand.Run(FillDirection.Right));
        public void OnFastFillDown(Office.IRibbonControl c) => Run("Fast Fill Down", () => FastFillCommand.Run(FillDirection.Down));

        // ---- Charts ----
        public void OnBuildWaterfall(Office.IRibbonControl c) => Run("Build Waterfall", () => BuildWaterfallCommand.Run());
        public void OnBuildStackedWaterfall(Office.IRibbonControl c) => Run("Build Stacked Waterfall", () => BuildStackedWaterfallCommand.Run());
        public void OnDisplayCagr(Office.IRibbonControl c) => Run("Display CAGR", () => CagrCommand.Run());

        // ---- Export / linking (W3) ----
        public void OnExportPpt(Office.IRibbonControl c) => Run("Export to PowerPoint", () => ExportToPowerPointCommand.Run(Core.Linking.ExportType.Picture));
        public void OnExportPptTable(Office.IRibbonControl c) => Run("Export as Table", () => ExportToPowerPointCommand.Run(Core.Linking.ExportType.Table));
        public void OnExportWord(Office.IRibbonControl c) => Run("Export to Word", () => ExportToWordCommand.Run(Core.Linking.ExportType.Picture));
        public void OnAdvancedExport(Office.IRibbonControl c) => Run("Advanced Export", () => AdvancedExportCommand.Run());
        public void OnHighlightLinked(Office.IRibbonControl c) => Run("Highlight Linked", () => HighlightLinkedCommand.Run());

        // ---- Library / Settings / Clean / Print (W5) ----
        public void OnSaveLibrary(Office.IRibbonControl c) => Run("Save to Library", () => ExcelLibraryCommand.SaveSelection());
        public void OnInsertLibrary(Office.IRibbonControl c) => Run("Insert from Library", () => ExcelLibraryCommand.InsertFromLibrary());
        public void OnCleanPrepare(Office.IRibbonControl c) => Run("Clean & Prepare", () => CleanPrepareCommand.Run());
        public void OnSmartPrint(Office.IRibbonControl c) => Run("Smart Print", () => SmartPrintCommand.Run());
        public void OnSettings(Office.IRibbonControl c) => Run("Settings", () => SettingsCommand.Run());

        // ---- Not yet implemented (later phases) ----
        public void OnNotImplemented(Office.IRibbonControl c)
        {
            MessageBox.Show("This feature lands in a later phase (see project plan §9).",
                "Upslide", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary>Wrap a command: log, run, surface a friendly message on failure (NFR-3).</summary>
        private void Run(string name, Func<string> action)
        {
            try
            {
                Logger.Info($"Command start: {name}");
                var status = action();
                Logger.Info($"Command ok: {name} — {status}");
            }
            catch (Exception ex)
            {
                Logger.Error($"Command failed: {name}", ex);
                MessageBox.Show(ex.Message, "Upslide — " + name,
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private static string ReadResource(string resourceName)
        {
            var asm = Assembly.GetExecutingAssembly();
            using (var stream = asm.GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                    throw new InvalidOperationException("Embedded ribbon XML not found: " + resourceName);
                using (var reader = new StreamReader(stream))
                    return reader.ReadToEnd();
            }
        }
    }
}
