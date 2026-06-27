using System;
using System.Collections.Generic;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using UpslideClone.Core.Charts;

namespace UpslideClone.Excel.Commands
{
    /// <summary>
    /// CAGR arrow (FR-C4). Reads the first series of the selected chart, computes
    /// the CAGR with <see cref="CagrEngine"/>, and overlays an arrow + label.
    /// W1 places the arrow diagonally across the plot area; exact data-point
    /// anchoring is a W-later refinement (the % is exact).
    /// </summary>
    internal static class CagrCommand
    {
        public static string Run()
        {
            var theme = ThemeProvider.Current;
            ExcelInterop.Chart chart = ExcelHelpers.App.ActiveChart;
            if (chart == null)
                throw new InvalidOperationException("Select a chart first, then run Display CAGR.");

            var seriesCol = (ExcelInterop.SeriesCollection)chart.SeriesCollection();
            if (seriesCol.Count < 1)
                throw new InvalidOperationException("The selected chart has no data series.");

            // Cast Item(1) to Series so .Values is a static call (avoids dynamic
            // dispatch, which mis-casts the COM SAFEARRAY Object[*] → Object[]).
            ExcelInterop.Series series1 = (ExcelInterop.Series)seriesCol.Item(1);
            var values = ReadDoubles((object)series1.Values);
            if (values.Count < 2)
                throw new InvalidOperationException("Need at least 2 data points to compute CAGR.");

            CagrResult cagr = CagrEngine.Compute(values);

            // Geometry of the plot area.
            ExcelInterop.PlotArea plot = chart.PlotArea;
            double x0 = plot.InsideLeft;
            double y0 = plot.InsideTop;
            double w = plot.InsideWidth;
            double h = plot.InsideHeight;

            // Arrow from the first point (lower-left) to the last point (upper-right)
            // for positive growth; reversed vertically for negative.
            double startX = x0 + w * 0.10;
            double endX = x0 + w * 0.90;
            double startY = cagr.Cagr >= 0 ? y0 + h * 0.85 : y0 + h * 0.20;
            double endY = cagr.Cagr >= 0 ? y0 + h * 0.20 : y0 + h * 0.85;

            ExcelInterop.Shape arrow = chart.Shapes.AddLine(
                (float)startX, (float)startY, (float)endX, (float)endY);
            arrow.Line.EndArrowheadStyle = Office.MsoArrowheadStyle.msoArrowheadTriangle;
            arrow.Line.Weight = 2.0f;
            arrow.Line.ForeColor.RGB = ExcelHelpers.Ole(theme.Colors.Total);

            ExcelInterop.Shape label = chart.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal,
                (float)(x0 + w * 0.40), (float)(y0 + h * 0.08), 120f, 22f);
            label.TextFrame.Characters().Text = cagr.Label;
            label.TextFrame.Characters().Font.Bold = true;
            label.TextFrame.Characters().Font.Name = theme.Fonts.Latin;
            label.Line.Visible = Office.MsoTriState.msoFalse;
            label.Fill.Visible = Office.MsoTriState.msoFalse;

            return cagr.Label + $" over {cagr.Periods} periods.";
        }

        private static List<double> ReadDoubles(object seriesValues)
        {
            var result = new List<double>();
            if (seriesValues is Array arr)
            {
                for (int i = arr.GetLowerBound(0); i <= arr.GetUpperBound(0); i++)
                {
                    object v = arr.GetValue(i);
                    if (v is double d && !double.IsNaN(d)) result.Add(d);
                    else if (v != null && double.TryParse(v.ToString(), out double parsed)) result.Add(parsed);
                }
            }
            return result;
        }
    }
}
