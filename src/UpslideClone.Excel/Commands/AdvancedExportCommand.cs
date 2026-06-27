using System;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using PPT = Microsoft.Office.Interop.PowerPoint;
using UpslideClone.Core.Linking;
using UpslideClone.Core.Util;

namespace UpslideClone.Excel.Commands
{
    /// <summary>
    /// Advanced Export (FR-X8): read the selected mapping table (range → slide),
    /// then batch-export every entry as a linked object in one re-runnable action.
    /// </summary>
    internal static class AdvancedExportCommand
    {
        public static string Run()
        {
            var sel = ExcelHelpers.RequireRangeSelection();
            var mapSheet = (ExcelInterop.Worksheet)sel.Worksheet;
            var book = (ExcelInterop.Workbook)mapSheet.Parent;

            var grid = sel.Value2 as object[,];
            var entries = AdvancedExportMap.Parse(grid);
            if (entries.Count == 0)
                throw new InvalidOperationException("No mapping rows found. Select a table with a header row (Range, Slide, Type, …) and at least one row.");

            var app = ExportToPowerPointCommand.GetOrStartPowerPoint();
            PPT.Presentation pres = ExportToPowerPointCommand.ActivePresentation(app);

            int ok = 0, failed = 0;
            foreach (var e in entries)
            {
                try
                {
                    ExcelInterop.Worksheet ws = string.IsNullOrEmpty(e.SourceSheet)
                        ? mapSheet
                        : (ExcelInterop.Worksheet)book.Worksheets[e.SourceSheet];
                    ExcelInterop.Range range = ws.Range[e.SourceRange];
                    ExportToPowerPointCommand.ExportRangeToSlide(app, pres, range, e.ExportType, e.TargetSlide, e.PlaceholderKey);
                    ok++;
                }
                catch (Exception ex)
                {
                    Logger.Error($"Advanced Export entry failed ({e.SourceSheet}!{e.SourceRange} → slide {e.TargetSlide})", ex);
                    failed++;
                }
            }

            try { app.Activate(); } catch { }
            return $"Advanced Export: {ok} object(s) linked" + (failed > 0 ? $", {failed} failed (see log)." : ".");
        }
    }
}
