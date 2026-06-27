using System;
using Wd = Microsoft.Office.Interop.Word;
using UpslideClone.Core.Branding;
using UpslideClone.Core.Util;

namespace UpslideClone.Word.Formatting
{
    /// <summary>
    /// Word branded-formatting toolkit (the Word analogue of Excel Smart Format):
    /// brand the selected table (green header, fonts, borders) and apply the brand
    /// font to a selection. Theme-driven via the shared <see cref="BrandTheme"/>.
    /// </summary>
    internal static class FormatCommands
    {
        public static string FormatTable(Wd.Application app)
        {
            var sel = app.Selection;
            if (sel.Tables.Count < 1)
                throw new InvalidOperationException("Place the cursor inside a table first.");

            var theme = BrandTheme.Default();
            var tbl = sel.Tables[1];

            // Whole-table font first, then header styling on top.
            tbl.Range.Font.Name = theme.Fonts.Latin;
            tbl.Range.Font.Size = (float)theme.Fonts.SizeBody;

            var header = tbl.Rows[1];
            header.Range.Shading.BackgroundPatternColor = (Wd.WdColor)ColorUtil.OleFromHex(theme.Colors.HeaderFill);
            header.Range.Font.Bold = 1;
            header.Range.Font.Color = (Wd.WdColor)ColorUtil.OleFromHex(theme.Colors.HeaderFont);
            header.Range.Font.Name = theme.Fonts.Latin;

            // Single borders in the brand grey.
            var borderColor = (Wd.WdColor)ColorUtil.OleFromHex(theme.Colors.Border);
            foreach (Wd.WdBorderType edge in new[]
            {
                Wd.WdBorderType.wdBorderTop, Wd.WdBorderType.wdBorderBottom,
                Wd.WdBorderType.wdBorderLeft, Wd.WdBorderType.wdBorderRight,
                Wd.WdBorderType.wdBorderHorizontal, Wd.WdBorderType.wdBorderVertical
            })
            {
                var b = tbl.Borders[edge];
                b.LineStyle = Wd.WdLineStyle.wdLineStyleSingle;
                b.Color = borderColor;
            }

            return $"Branded table ({tbl.Rows.Count}×{tbl.Columns.Count}).";
        }

        public static string ApplyBrandFont(Wd.Application app)
        {
            var theme = BrandTheme.Default();
            var sel = app.Selection;
            if (sel == null) throw new InvalidOperationException("Select some text first.");
            sel.Font.Name = theme.Fonts.Latin;
            return "Applied brand font to the selection.";
        }
    }
}
