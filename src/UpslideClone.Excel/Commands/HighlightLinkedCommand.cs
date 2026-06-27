using System;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using UpslideClone.Core.Linking;

namespace UpslideClone.Excel.Commands
{
    /// <summary>
    /// Highlight linked items (FR-X10): flag every source range in this workbook
    /// that has been exported/linked out, using the CustomXMLPart registry.
    /// Re-running clears prior highlights first.
    /// </summary>
    internal static class HighlightLinkedCommand
    {
        // A distinctive marker fill (brand green tint) applied to linked source ranges.
        private const string HighlightHex = "#E2EFDA";

        public static string Run()
        {
            var app = ExcelHelpers.App;
            var book = (ExcelInterop.Workbook)app.ActiveWorkbook;
            if (book == null) throw new InvalidOperationException("Open a workbook first.");

            var reg = LinkRegistryStore.Load(book);
            if (reg.Items.Count == 0)
                return "No linked ranges recorded in this workbook yet.";

            int ole = ExcelHelpers.Ole(HighlightHex);
            int highlighted = 0, missing = 0;
            foreach (var item in reg.Items)
            {
                try
                {
                    var ws = (ExcelInterop.Worksheet)book.Worksheets[item.Sheet];
                    ExcelInterop.Range range = ws.Range[item.Range];
                    range.Interior.Color = ole;
                    range.BorderAround(ExcelInterop.XlLineStyle.xlContinuous, ExcelInterop.XlBorderWeight.xlMedium,
                        ExcelInterop.XlColorIndex.xlColorIndexAutomatic, missingArg);
                    highlighted++;
                }
                catch
                {
                    missing++; // sheet/range no longer exists
                }
            }

            return $"Highlighted {highlighted} linked range(s)."
                 + (missing > 0 ? $" {missing} could not be resolved." : "");
        }

        private static readonly object missingArg = Type.Missing;
    }
}
