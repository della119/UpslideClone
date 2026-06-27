using System;
using System.Collections.Generic;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace UpslideClone.Excel.Commands
{
    /// <summary>
    /// In-memory undo cache for Smart Format (FR-F5). Before a format action, the
    /// affected cells' prior style is snapshotted; "Undo Formatting" replays the
    /// most recent snapshot cell-by-cell. Values are never touched.
    ///
    /// W1 captures cell-level style (number format, font, fill, alignment) and
    /// resets borders on undo. Per-edge border capture is a W2 refinement.
    /// </summary>
    internal static class UndoCache
    {
        private sealed class CellStyle
        {
            public string NumberFormat;
            public string FontName;
            public double FontSize;
            public bool FontBold;
            public int FontColor;
            public object InteriorColor; // may be int or Missing (no fill)
            public int HAlign;
        }

        private sealed class Snapshot
        {
            public string SheetName;
            public int Row;     // 1-based top-left
            public int Col;
            public int Rows;
            public int Cols;
            public CellStyle[,] Cells;
        }

        private static readonly Stack<Snapshot> History = new Stack<Snapshot>();

        public static bool HasHistory => History.Count > 0;

        /// <summary>Capture the current style of <paramref name="range"/> onto the undo stack.</summary>
        public static void Capture(ExcelInterop.Range range)
        {
            var sheet = (ExcelInterop.Worksheet)range.Worksheet;
            int rows = range.Rows.Count;
            int cols = range.Columns.Count;
            int top = range.Row;
            int left = range.Column;

            var snap = new Snapshot
            {
                SheetName = sheet.Name,
                Row = top,
                Col = left,
                Rows = rows,
                Cols = cols,
                Cells = new CellStyle[rows, cols]
            };

            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    var cell = (ExcelInterop.Range)sheet.Cells[top + r, left + c];
                    var font = cell.Font;
                    var interior = cell.Interior;
                    snap.Cells[r, c] = new CellStyle
                    {
                        NumberFormat = cell.NumberFormat as string,
                        FontName = font.Name as string,
                        FontSize = Convert.ToDouble(font.Size),
                        FontBold = Convert.ToBoolean(font.Bold),
                        FontColor = Convert.ToInt32(font.Color),
                        InteriorColor = interior.Pattern is int p && p == (int)ExcelInterop.XlPattern.xlPatternNone
                            ? (object)null
                            : interior.Color,
                        HAlign = Convert.ToInt32(cell.HorizontalAlignment)
                    };
                }
            }

            History.Push(snap);
        }

        /// <summary>Restore the most recent snapshot; returns a status string.</summary>
        public static string UndoLast()
        {
            if (History.Count == 0)
                throw new InvalidOperationException("Nothing to undo.");

            var snap = History.Pop();
            var sheet = (ExcelInterop.Worksheet)ExcelHelpers.App.Sheets[snap.SheetName];

            for (int r = 0; r < snap.Rows; r++)
            {
                for (int c = 0; c < snap.Cols; c++)
                {
                    var cs = snap.Cells[r, c];
                    var cell = (ExcelInterop.Range)sheet.Cells[snap.Row + r, snap.Col + c];
                    cell.NumberFormat = cs.NumberFormat ?? "General";
                    cell.Font.Name = cs.FontName;
                    cell.Font.Size = cs.FontSize;
                    cell.Font.Bold = cs.FontBold;
                    cell.Font.Color = cs.FontColor;
                    if (cs.InteriorColor == null)
                        cell.Interior.Pattern = ExcelInterop.XlPattern.xlPatternNone;
                    else
                        cell.Interior.Color = cs.InteriorColor;
                    cell.HorizontalAlignment = cs.HAlign;
                }
            }

            // Reset borders over the block (W1 limitation: per-edge capture is W2).
            var block = ExcelHelpers.Sub(sheet, snap.Row, snap.Col, snap.Rows, snap.Cols);
            block.Borders.LineStyle = ExcelInterop.XlLineStyle.xlLineStyleNone;

            return $"Reverted formatting on {snap.Rows}×{snap.Cols} cells.";
        }

        public static void Clear() => History.Clear();
    }
}
