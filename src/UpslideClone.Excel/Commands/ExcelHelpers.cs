using System;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using UpslideClone.Core.Util;

namespace UpslideClone.Excel.Commands
{
    /// <summary>Shared Excel-Interop helpers for the W1 commands.</summary>
    internal static class ExcelHelpers
    {
        public static ExcelInterop.Application App => Globals.ThisAddIn.Application;

        /// <summary>The current selection as a Range, or throw a friendly error.</summary>
        public static ExcelInterop.Range RequireRangeSelection()
        {
            var sel = App.Selection as ExcelInterop.Range;
            if (sel == null)
                throw new InvalidOperationException("Select a cell range first.");
            return sel;
        }

        /// <summary>Absolute sub-range by 1-based row/col offset from a worksheet origin.</summary>
        public static ExcelInterop.Range Sub(ExcelInterop.Worksheet sheet, int row, int col, int rows, int cols)
        {
            return sheet.Range[sheet.Cells[row, col], sheet.Cells[row + rows - 1, col + cols - 1]];
        }

        /// <summary>BGR/OLE integer from a "#RRGGBB" hex string (Interop expects BGR).</summary>
        public static int Ole(string hex) => ColorUtil.OleFromHex(hex);

        /// <summary>Best-effort COM release in finally blocks.</summary>
        public static void Release(object com)
        {
            try
            {
                if (com != null && System.Runtime.InteropServices.Marshal.IsComObject(com))
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(com);
            }
            catch { /* never throw from cleanup */ }
        }
    }
}
