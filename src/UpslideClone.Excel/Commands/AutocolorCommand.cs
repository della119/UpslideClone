using System;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using UpslideClone.Core.Modelling;

namespace UpslideClone.Excel.Commands
{
    /// <summary>
    /// Autocolor (FR-M1) — applies the banker colour convention across the
    /// selection using the pure <see cref="AutocolorClassifier"/>:
    /// blue = inputs, black = formulas, green = links/external refs.
    /// </summary>
    internal static class AutocolorCommand
    {
        public static string Run()
        {
            var sel = ExcelHelpers.RequireRangeSelection();
            UndoCache.Capture(sel);

            int colored = 0;
            // Iterate only cells with content or formulas; ignore blanks.
            foreach (ExcelInterop.Range cell in sel.Cells)
            {
                bool hasFormula = Convert.ToBoolean(cell.HasFormula);
                bool isEmpty = cell.Value2 == null && !hasFormula;
                string formula = hasFormula ? (cell.Formula as string) : null;

                var cls = AutocolorClassifier.Classify(hasFormula, formula, isEmpty);
                if (cls == CellColorClass.Empty) continue;

                cell.Font.Color = ExcelHelpers.Ole(AutocolorClassifier.DefaultHex(cls));
                colored++;
            }

            return $"Autocolored {colored} cells.";
        }
    }
}
