using System;
using System.Collections.Generic;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using UpslideClone.Core.Charts;

namespace UpslideClone.Excel.Commands
{
    /// <summary>
    /// Stacked Waterfall builder (FR-C2). Reads a label column + N category
    /// columns (e.g. FR/UK/DE), computes the shared Base float plus per-category
    /// segments via <see cref="StackedWaterfallEngine"/>, writes a helper block,
    /// and builds a stacked column chart with an invisible Base series.
    /// </summary>
    internal static class BuildStackedWaterfallCommand
    {
        public static string Run()
        {
            var theme = ThemeProvider.Current;
            var sel = ExcelHelpers.RequireRangeSelection();
            var sheet = (ExcelInterop.Worksheet)sel.Worksheet;

            int rows = sel.Rows.Count;
            int cols = sel.Columns.Count;
            if (rows < 3 || cols < 3)
                throw new InvalidOperationException("Select a header row, a label column, and at least 2 category columns.");

            var vals = (object[,])sel.Value2; // 1-based

            int nCats = cols - 1;
            var categories = new string[nCats];
            for (int c = 0; c < nCats; c++)
                categories[c] = vals[1, c + 2]?.ToString() ?? ("Cat" + (c + 1));

            var points = new List<StackedWaterfallPoint>();
            for (int r = 2; r <= rows; r++)
            {
                string label = vals[r, 1]?.ToString() ?? "";
                var rowVals = new double[nCats];
                bool anyNumeric = false;
                for (int c = 0; c < nCats; c++)
                {
                    if (vals[r, c + 2] is double d) { rowVals[c] = d; anyNumeric = true; }
                    else rowVals[c] = 0;
                }
                if (anyNumeric) points.Add(new StackedWaterfallPoint(label, rowVals));
            }
            if (points.Count < 2)
                throw new InvalidOperationException("Need at least 2 data rows to build a stacked waterfall.");

            var res = StackedWaterfallEngine.Compute(categories, points);

            // Helper block: Item | Base | cat1 | cat2 | ...
            int blockCols = 2 + nCats;
            var block = new object[res.Rows.Count + 1, blockCols];
            block[0, 0] = "Item";
            block[0, 1] = "Base";
            for (int c = 0; c < nCats; c++) block[0, c + 2] = categories[c];
            for (int i = 0; i < res.Rows.Count; i++)
            {
                var w = res.Rows[i];
                block[i + 1, 0] = w.Label;
                block[i + 1, 1] = w.Base;
                for (int c = 0; c < nCats; c++)
                    block[i + 1, c + 2] = (object)w.Segments[c] ?? "=NA()";
            }

            int startRow = sel.Row + rows + 2;
            int startCol = sel.Column;
            var helper = ExcelHelpers.Sub(sheet, startRow, startCol, res.Rows.Count + 1, blockCols);
            helper.Formula = block;
            helper.Font.Name = theme.Fonts.Latin;
            helper.Font.Size = 9;

            ExcelInterop.Shape shp = sheet.Shapes.AddChart2(-1, ExcelInterop.XlChartType.xlColumnStacked);
            ExcelInterop.Chart chart = shp.Chart;
            chart.SetSourceData(helper, ExcelInterop.XlRowCol.xlColumns);

            string[] palette = { theme.Colors.Increase, theme.Colors.Total, theme.Colors.Decrease };
            var seriesCol = (ExcelInterop.SeriesCollection)chart.SeriesCollection();
            int catIdx = 0;
            for (int i = 1; i <= seriesCol.Count; i++)
            {
                ExcelInterop.Series s = seriesCol.Item(i);
                if (s.Name == "Base")
                {
                    s.Format.Fill.Visible = Office.MsoTriState.msoFalse;
                }
                else
                {
                    s.Format.Fill.ForeColor.RGB = ExcelHelpers.Ole(palette[catIdx % palette.Length]);
                    catIdx++;
                }
            }

            chart.HasTitle = true;
            chart.ChartTitle.Text = "Stacked Waterfall";
            chart.HasLegend = true;
            chart.Legend.Position = ExcelInterop.XlLegendPosition.xlLegendPositionBottom;

            // Hide the Base legend entry.
            try
            {
                for (int i = 1; i <= seriesCol.Count; i++)
                    if (((ExcelInterop.Series)seriesCol.Item(i)).Name == "Base") { chart.Legend.LegendEntries(i).Delete(); break; }
            }
            catch { }

            var topCell = (ExcelInterop.Range)sheet.Cells[startRow, startCol];
            var leftCell = (ExcelInterop.Range)sheet.Cells[startRow, startCol + blockCols + 1];
            shp.Top = (float)topCell.Top + 10f;
            shp.Left = (float)leftCell.Left;

            return $"Stacked waterfall built: {res.Rows.Count} steps × {nCats} categories.";
        }
    }
}
