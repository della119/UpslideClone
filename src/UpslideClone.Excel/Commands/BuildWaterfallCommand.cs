using System;
using System.Collections.Generic;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using UpslideClone.Core.Charts;

namespace UpslideClone.Excel.Commands
{
    /// <summary>
    /// Waterfall builder — Interop port of addin/src/core/waterfall.ts and
    /// Appendix A.2. Writes a Base/Decrease/Increase/Total helper block
    /// (=NA() for non-applicable series) below the selection, then builds a
    /// stacked column chart with an invisible Base series.
    /// </summary>
    internal static class BuildWaterfallCommand
    {
        public static string Run()
        {
            var theme = ThemeProvider.Current;
            var sel = ExcelHelpers.RequireRangeSelection();
            var sheet = (ExcelInterop.Worksheet)sel.Worksheet;

            int rows = sel.Rows.Count;
            int cols = sel.Columns.Count;
            if (cols < 2)
                throw new InvalidOperationException("Select a two-column table: labels in the left column, values in the next.");

            var vals = (object[,])sel.Value2; // 1-based

            // Parse label/value pairs; skip non-numeric (title/header) rows.
            var points = new List<WaterfallPoint>();
            for (int r = 1; r <= rows; r++)
            {
                string label = vals[r, 1]?.ToString() ?? "";
                if (vals[r, 2] is double d && !double.IsNaN(d))
                    points.Add(new WaterfallPoint(label, d));
            }
            if (points.Count < 2)
                throw new InvalidOperationException("Need at least 2 numeric rows to build a waterfall.");

            var wfRows = WaterfallEngine.Compute(points);

            // Build the helper block (header + rows × 5).
            var block = new object[wfRows.Count + 1, 5];
            block[0, 0] = "Item"; block[0, 1] = "Base"; block[0, 2] = "Decrease"; block[0, 3] = "Increase"; block[0, 4] = "Total";
            for (int i = 0; i < wfRows.Count; i++)
            {
                var w = wfRows[i];
                block[i + 1, 0] = w.Label;
                block[i + 1, 1] = w.Base;
                block[i + 1, 2] = (object)w.Decrease ?? "=NA()";
                block[i + 1, 3] = (object)w.Increase ?? "=NA()";
                block[i + 1, 4] = (object)w.Total ?? "=NA()";
            }

            int startRow = sel.Row + rows + 2;
            int startCol = sel.Column;
            var helper = ExcelHelpers.Sub(sheet, startRow, startCol, wfRows.Count + 1, 5);
            helper.Formula = block; // numbers stay numbers; "=NA()" becomes #N/A
            helper.Font.Name = theme.Fonts.Latin;
            helper.Font.Size = 9;

            // Build the stacked column chart from the helper block.
            ExcelInterop.Shape shp = sheet.Shapes.AddChart2(-1, ExcelInterop.XlChartType.xlColumnStacked);
            ExcelInterop.Chart chart = shp.Chart;
            chart.SetSourceData(helper, ExcelInterop.XlRowCol.xlColumns);

            var seriesCol = (ExcelInterop.SeriesCollection)chart.SeriesCollection();
            for (int i = 1; i <= seriesCol.Count; i++)
            {
                ExcelInterop.Series s = seriesCol.Item(i);
                switch (s.Name)
                {
                    case "Base":
                        s.Format.Fill.Visible = Office.MsoTriState.msoFalse; // invisible floor
                        break;
                    case "Increase":
                        s.Format.Fill.ForeColor.RGB = ExcelHelpers.Ole(theme.Colors.Increase);
                        SetLabels(s, theme.NumberFormats.Number);
                        break;
                    case "Decrease":
                        s.Format.Fill.ForeColor.RGB = ExcelHelpers.Ole(theme.Colors.Decrease);
                        SetLabels(s, theme.NumberFormats.Number);
                        break;
                    case "Total":
                        s.Format.Fill.ForeColor.RGB = ExcelHelpers.Ole(theme.Colors.Total);
                        SetLabels(s, theme.NumberFormats.Number);
                        break;
                }
            }

            chart.HasTitle = true;
            chart.ChartTitle.Text = "Waterfall";
            chart.HasLegend = true;
            chart.Legend.Position = ExcelInterop.XlLegendPosition.xlLegendPositionBottom;

            // Hide the "Base" entry from the legend.
            try
            {
                for (int i = 1; i <= seriesCol.Count; i++)
                {
                    if (((ExcelInterop.Series)seriesCol.Item(i)).Name == "Base")
                    {
                        chart.Legend.LegendEntries(i).Delete();
                        break;
                    }
                }
            }
            catch { /* legend entry API differences — leave as is */ }

            var topCell = (ExcelInterop.Range)sheet.Cells[startRow, startCol];
            var leftCell = (ExcelInterop.Range)sheet.Cells[startRow, startCol + 6];
            shp.Top = (float)topCell.Top + 10f;
            shp.Left = (float)leftCell.Left;

            return $"Waterfall built from {points.Count} points.";
        }

        private static void SetLabels(ExcelInterop.Series s, string numberFormat)
        {
            try
            {
                s.HasDataLabels = true;
                ((ExcelInterop.DataLabels)s.DataLabels()).NumberFormat = numberFormat;
            }
            catch { /* older API: skip per-series labels */ }
        }
    }
}
