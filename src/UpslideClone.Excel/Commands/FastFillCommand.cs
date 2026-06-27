using System;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using UpslideClone.Core.Modelling;

namespace UpslideClone.Excel.Commands
{
    internal enum FillDirection { Right, Down }

    /// <summary>
    /// Fast Fill right/down (FR-M3). Takes the top-left cell's formula and fills it
    /// across the selection, computing each target with the verified
    /// <see cref="FormulaReferences.Shift"/> so relative refs move and $-anchors
    /// stay — structure-aware, without relying on Excel's autofill heuristics.
    /// </summary>
    internal static class FastFillCommand
    {
        public static string Run(FillDirection direction)
        {
            var sel = ExcelHelpers.RequireRangeSelection();
            var sheet = (ExcelInterop.Worksheet)sel.Worksheet;
            int rows = sel.Rows.Count;
            int cols = sel.Columns.Count;
            int top = sel.Row;
            int left = sel.Column;

            var source = (ExcelInterop.Range)sheet.Cells[top, left];
            if (!Convert.ToBoolean(source.HasFormula))
                throw new InvalidOperationException("The top-left cell of the selection has no formula to fill.");

            string srcFormula = source.Formula as string;
            UndoCache.Capture(sel);

            int filled = 0;
            if (direction == FillDirection.Right)
            {
                for (int c = 1; c < cols; c++)
                {
                    var target = (ExcelInterop.Range)sheet.Cells[top, left + c];
                    target.Formula = FormulaReferences.Shift(srcFormula, 0, c);
                    filled++;
                }
            }
            else
            {
                for (int r = 1; r < rows; r++)
                {
                    var target = (ExcelInterop.Range)sheet.Cells[top + r, left];
                    target.Formula = FormulaReferences.Shift(srcFormula, r, 0);
                    filled++;
                }
            }

            return $"Fast-filled {filled} cell(s) {direction.ToString().ToLowerInvariant()}.";
        }
    }
}
