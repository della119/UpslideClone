using System;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace UpslideClone.Excel.Commands
{
    /// <summary>
    /// Smart Print (FR-M7): apply standardized print settings to the active sheet
    /// (landscape, fit-to-one-page-wide, narrow margins, centered) and show a
    /// print preview. Consistent output without manual page setup.
    /// </summary>
    internal static class SmartPrintCommand
    {
        public static string Run()
        {
            var app = ExcelHelpers.App;
            var ws = app.ActiveSheet as ExcelInterop.Worksheet;
            if (ws == null) throw new InvalidOperationException("Select a worksheet first.");

            var ps = ws.PageSetup;
            ps.Orientation = ExcelInterop.XlPageOrientation.xlLandscape;
            ps.Zoom = false;
            ps.FitToPagesWide = 1;
            ps.FitToPagesTall = false; // as many pages tall as needed
            ps.CenterHorizontally = true;
            ps.LeftMargin = app.InchesToPoints(0.5);
            ps.RightMargin = app.InchesToPoints(0.5);
            ps.TopMargin = app.InchesToPoints(0.75);
            ps.BottomMargin = app.InchesToPoints(0.75);
            ps.PrintGridlines = false;

            try { ws.PrintPreview(); } catch { /* preview unavailable in some contexts */ }
            return "Applied standardized print settings (landscape, fit-to-width).";
        }
    }
}
