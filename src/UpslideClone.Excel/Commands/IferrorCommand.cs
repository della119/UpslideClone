using System;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using UpslideClone.Core.Modelling;

namespace UpslideClone.Excel.Commands
{
    /// <summary>
    /// IFERROR wrapper (FR-M2) — wraps every formula cell in the selection in
    /// IFERROR(…,"") via the pure <see cref="FormulaTransform"/>. Constants and
    /// already-wrapped formulas are left untouched.
    /// </summary>
    internal static class IferrorCommand
    {
        public static string Run(string replacement = "\"\"")
        {
            var sel = ExcelHelpers.RequireRangeSelection();

            int wrapped = 0;
            foreach (ExcelInterop.Range cell in sel.Cells)
            {
                if (!Convert.ToBoolean(cell.HasFormula)) continue;

                string formula = cell.Formula as string;
                string updated = FormulaTransform.WrapIfError(formula, replacement);
                if (!string.Equals(updated, formula, StringComparison.Ordinal))
                {
                    cell.Formula = updated;
                    wrapped++;
                }
            }

            return $"Wrapped {wrapped} formula(s) in IFERROR.";
        }
    }
}
