using System;
using System.Collections.Generic;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using UpslideClone.Core.Branding;
using UpslideClone.Core.Formatting;

namespace UpslideClone.Excel.Commands
{
    internal enum SmartFormatDimension
    {
        Title,
        Result,
        NumberFormat,
        PercentFormat
    }

    /// <summary>
    /// Smart Format (tables) — Interop port of addin/src/core/smartFormat.ts and
    /// Appendix A.3. Header row, auto-detected result/total rows, per-column number
    /// vs percent formats, alignment, and outer border, all theme-driven.
    /// </summary>
    internal static class SmartFormatCommand
    {
        public static string Run()
        {
            var theme = ThemeProvider.Current;
            var sel = ExcelHelpers.RequireRangeSelection();
            var sheet = (ExcelInterop.Worksheet)sel.Worksheet;

            int rows = sel.Rows.Count;
            int cols = sel.Columns.Count;
            if (rows < 2 || cols < 2)
                throw new InvalidOperationException("Select a table with at least 2 rows and 2 columns.");

            UndoCache.Capture(sel);

            int top = sel.Row;
            int left = sel.Column;
            var vals = (object[,])sel.Value2; // 1-based [row, col]

            // Decide percent vs number per data column.
            var pctCol = new bool[cols + 1];
            for (int c = 2; c <= cols; c++)
            {
                string header = AsString(vals[1, c]);
                var colVals = new List<double>();
                for (int r = 2; r <= rows; r++)
                    if (vals[r, c] is double d) colVals.Add(d);
                pctCol[c] = SmartFormatRules.IsPercentColumn(header, colVals);
            }

            // Base font over the whole range.
            sel.Font.Name = theme.Fonts.Latin;
            sel.Font.Size = theme.Fonts.SizeBody;

            // Number formats per data column body.
            for (int c = 2; c <= cols; c++)
            {
                var body = ExcelHelpers.Sub(sheet, top + 1, left + (c - 1), rows - 1, 1);
                body.NumberFormat = pctCol[c] ? theme.NumberFormats.Percent : theme.NumberFormats.Number;
            }

            // Alignment: labels left, numbers right.
            ExcelHelpers.Sub(sheet, top + 1, left, rows - 1, 1).HorizontalAlignment =
                ExcelInterop.XlHAlign.xlHAlignLeft;
            ExcelHelpers.Sub(sheet, top + 1, left + 1, rows - 1, cols - 1).HorizontalAlignment =
                ExcelInterop.XlHAlign.xlHAlignRight;

            ApplyHeader(sheet, top, left, cols, theme, on: true);
            ApplyResultRows(sheet, vals, top, left, rows, cols, theme, on: true);

            // Outer border.
            ApplyOuterBorder(sel, theme);

            return $"Formatted {rows}×{cols} table.";
        }

        public static string ClearFormatting()
        {
            var sel = ExcelHelpers.RequireRangeSelection();
            UndoCache.Capture(sel);
            sel.ClearFormats();
            return "Cleared formatting.";
        }

        /// <summary>Flip a single style dimension on/off (FR-F4), capturing undo state first.</summary>
        public static string Toggle(SmartFormatDimension dim)
        {
            var theme = ThemeProvider.Current;
            var sel = ExcelHelpers.RequireRangeSelection();
            var sheet = (ExcelInterop.Worksheet)sel.Worksheet;
            int rows = sel.Rows.Count, cols = sel.Columns.Count;
            if (rows < 2 || cols < 2)
                throw new InvalidOperationException("Select a table with at least 2 rows and 2 columns.");

            UndoCache.Capture(sel);
            int top = sel.Row, left = sel.Column;
            var vals = (object[,])sel.Value2;

            switch (dim)
            {
                case SmartFormatDimension.Title:
                {
                    bool on = !HeaderIsStyled(sheet, top, left, theme);
                    ApplyHeader(sheet, top, left, cols, theme, on);
                    return on ? "Title formatting on." : "Title formatting off.";
                }
                case SmartFormatDimension.Result:
                {
                    bool on = !AnyResultRowStyled(sheet, vals, top, left, rows, cols, theme);
                    ApplyResultRows(sheet, vals, top, left, rows, cols, theme, on);
                    return on ? "Result formatting on." : "Result formatting off.";
                }
                case SmartFormatDimension.NumberFormat:
                case SmartFormatDimension.PercentFormat:
                {
                    bool percentOnly = dim == SmartFormatDimension.PercentFormat;
                    var probe = (ExcelInterop.Range)sheet.Cells[top + 1, left + 1];
                    bool on = string.Equals(probe.NumberFormat as string, "General", StringComparison.OrdinalIgnoreCase);
                    for (int c = 2; c <= cols; c++)
                    {
                        string header = AsString(vals[1, c]);
                        var colVals = new List<double>();
                        for (int r = 2; r <= rows; r++)
                            if (vals[r, c] is double d) colVals.Add(d);
                        bool isPct = SmartFormatRules.IsPercentColumn(header, colVals);
                        if (percentOnly && !isPct) continue;

                        var body = ExcelHelpers.Sub(sheet, top + 1, left + (c - 1), rows - 1, 1);
                        body.NumberFormat = on
                            ? (isPct ? theme.NumberFormats.Percent : theme.NumberFormats.Number)
                            : "General";
                    }
                    return on ? "Number formats on." : "Number formats off.";
                }
            }
            return "No-op.";
        }

        // ---- dimension appliers ----

        private static void ApplyHeader(ExcelInterop.Worksheet sheet, int top, int left, int cols, BrandTheme theme, bool on)
        {
            var header = ExcelHelpers.Sub(sheet, top, left, 1, cols);
            if (on)
            {
                header.Interior.Color = ExcelHelpers.Ole(theme.Colors.HeaderFill);
                header.Font.Color = ExcelHelpers.Ole(theme.Colors.HeaderFont);
                header.Font.Bold = true;
                header.Borders[ExcelInterop.XlBordersIndex.xlEdgeBottom].LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
            }
            else
            {
                header.Interior.Pattern = ExcelInterop.XlPattern.xlPatternNone;
                header.Font.ColorIndex = ExcelInterop.XlColorIndex.xlColorIndexAutomatic;
                header.Font.Bold = false;
                header.Borders[ExcelInterop.XlBordersIndex.xlEdgeBottom].LineStyle = ExcelInterop.XlLineStyle.xlLineStyleNone;
            }
        }

        private static void ApplyResultRows(ExcelInterop.Worksheet sheet, object[,] vals, int top, int left, int rows, int cols, BrandTheme theme, bool on)
        {
            for (int r = 2; r <= rows; r++)
            {
                string label = AsString(vals[r, 1]);
                if (!SmartFormatRules.IsResultRow(label, theme.ResultRowKeywords)) continue;

                var rng = ExcelHelpers.Sub(sheet, top + (r - 1), left, 1, cols);
                if (on)
                {
                    rng.Font.Bold = true;
                    rng.Interior.Color = ExcelHelpers.Ole(theme.Colors.ResultFill);
                    rng.Borders[ExcelInterop.XlBordersIndex.xlEdgeTop].LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
                }
                else
                {
                    rng.Font.Bold = false;
                    rng.Interior.Pattern = ExcelInterop.XlPattern.xlPatternNone;
                    rng.Borders[ExcelInterop.XlBordersIndex.xlEdgeTop].LineStyle = ExcelInterop.XlLineStyle.xlLineStyleNone;
                }
            }
        }

        private static void ApplyOuterBorder(ExcelInterop.Range sel, BrandTheme theme)
        {
            int ole = ExcelHelpers.Ole(theme.Colors.Border);
            foreach (ExcelInterop.XlBordersIndex edge in new[]
            {
                ExcelInterop.XlBordersIndex.xlEdgeTop,
                ExcelInterop.XlBordersIndex.xlEdgeBottom,
                ExcelInterop.XlBordersIndex.xlEdgeLeft,
                ExcelInterop.XlBordersIndex.xlEdgeRight
            })
            {
                var b = sel.Borders[edge];
                b.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
                b.Color = ole;
            }
        }

        // ---- state probes (for toggles) ----

        private static bool HeaderIsStyled(ExcelInterop.Worksheet sheet, int top, int left, BrandTheme theme)
        {
            var cell = (ExcelInterop.Range)sheet.Cells[top, left];
            if (cell.Interior.Pattern is int p && p == (int)ExcelInterop.XlPattern.xlPatternNone) return false;
            try { return Convert.ToInt32(cell.Interior.Color) == ExcelHelpers.Ole(theme.Colors.HeaderFill); }
            catch { return false; }
        }

        private static bool AnyResultRowStyled(ExcelInterop.Worksheet sheet, object[,] vals, int top, int left, int rows, int cols, BrandTheme theme)
        {
            for (int r = 2; r <= rows; r++)
            {
                string label = AsString(vals[r, 1]);
                if (!SmartFormatRules.IsResultRow(label, theme.ResultRowKeywords)) continue;
                var cell = (ExcelInterop.Range)sheet.Cells[top + (r - 1), left];
                if (Convert.ToBoolean(cell.Font.Bold)) return true;
            }
            return false;
        }

        private static string AsString(object o) => o == null ? "" : o.ToString();
    }
}
