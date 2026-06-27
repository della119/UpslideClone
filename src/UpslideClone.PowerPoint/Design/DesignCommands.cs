using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PPT = Microsoft.Office.Interop.PowerPoint;
using UpslideClone.Core.Branding;
using UpslideClone.Core.Design;
using UpslideClone.Core.Util;

namespace UpslideClone.PowerPoint.Design
{
    /// <summary>
    /// PowerPoint design-toolkit commands (Smart Align, Resize &amp; Distribute,
    /// Arrange, Format, Select Similar, Smart Painter, Table of Contents, Slide
    /// Check). Geometry is computed by the tested <see cref="AlignEngine"/>; these
    /// are the thin Interop shells.
    /// </summary>
    internal static class DesignCommands
    {
        private static PPT.Application App => Globals.ThisAddIn.Application;

        private static PPT.ShapeRange RequireShapes(int min = 1)
        {
            var sel = App.ActiveWindow.Selection;
            if (sel.Type != PPT.PpSelectionType.ppSelectionShapes)
                throw new InvalidOperationException("Select one or more shapes first.");
            var sr = sel.ShapeRange;
            if (sr.Count < min)
                throw new InvalidOperationException($"Select at least {min} shapes.");
            return sr;
        }

        private static List<LayoutBox> ReadBoxes(PPT.ShapeRange sr)
        {
            var boxes = new List<LayoutBox>(sr.Count);
            for (int i = 1; i <= sr.Count; i++)
            {
                var s = sr[i];
                boxes.Add(new LayoutBox(s.Left, s.Top, s.Width, s.Height));
            }
            return boxes;
        }

        // ---- Smart Align (8) ----
        public static string Align(AlignMode mode)
        {
            var sr = RequireShapes(2);
            var result = AlignEngine.Align(ReadBoxes(sr), mode);
            for (int i = 1; i <= sr.Count; i++) { sr[i].Left = result[i - 1].Left; sr[i].Top = result[i - 1].Top; }
            return $"Aligned {sr.Count} shapes ({mode}).";
        }

        // ---- Resize & Distribute (9) ----
        public static string Distribute(DistributeAxis axis)
        {
            var sr = RequireShapes(3);
            var result = AlignEngine.Distribute(ReadBoxes(sr), axis);
            for (int i = 1; i <= sr.Count; i++) { sr[i].Left = result[i - 1].Left; sr[i].Top = result[i - 1].Top; }
            return $"Distributed {sr.Count} shapes ({axis}).";
        }

        public static string MatchSize(SizeMatch which)
        {
            var sr = RequireShapes(2);
            var result = AlignEngine.MatchSize(ReadBoxes(sr), which);
            for (int i = 1; i <= sr.Count; i++) { sr[i].Width = result[i - 1].Width; sr[i].Height = result[i - 1].Height; }
            return $"Matched size of {sr.Count} shapes ({which}).";
        }

        // ---- Arrange (5) ----
        public static string ZOrder(bool front)
        {
            var sr = RequireShapes(1);
            sr.ZOrder(front ? Office.MsoZOrderCmd.msoBringToFront : Office.MsoZOrderCmd.msoSendToBack);
            return front ? "Brought to front." : "Sent to back.";
        }

        public static string Group(bool group)
        {
            var sr = RequireShapes(1);
            if (group) { sr.Group(); return "Grouped."; }
            sr.Ungroup(); return "Ungrouped.";
        }

        // ---- Format shapes, text & tables (6) ----
        public static string FormatShapes()
        {
            var theme = BrandTheme.Default();
            int headerOle = ColorUtil.OleFromHex(theme.Colors.HeaderFill);
            int headerFontOle = ColorUtil.OleFromHex(theme.Colors.HeaderFont);
            var sr = RequireShapes(1);
            int n = 0;

            for (int i = 1; i <= sr.Count; i++)
            {
                var s = sr[i];
                try
                {
                    if (s.HasTable == Office.MsoTriState.msoTrue)
                    {
                        var tbl = s.Table;
                        for (int c = 1; c <= tbl.Columns.Count; c++)
                        {
                            var cell = tbl.Cell(1, c);
                            cell.Shape.Fill.ForeColor.RGB = headerOle;
                            var tr = cell.Shape.TextFrame.TextRange;
                            tr.Font.Bold = Office.MsoTriState.msoTrue;
                            tr.Font.Color.RGB = headerFontOle;
                            tr.Font.Name = theme.Fonts.Latin;
                        }
                        n++;
                    }
                    else if (s.HasTextFrame == Office.MsoTriState.msoTrue && s.TextFrame.HasText == Office.MsoTriState.msoTrue)
                    {
                        s.TextFrame.TextRange.Font.Name = theme.Fonts.Latin;
                        n++;
                    }
                }
                catch { /* skip shapes that don't support it */ }
            }
            return $"Applied branding to {n} shape(s).";
        }

        // ---- Select Similar (7) ----
        public static string SelectSimilar()
        {
            var sr = RequireShapes(1);
            var refShape = sr[1];
            var slide = (PPT.Slide)refShape.Parent;
            int refType = (int)refShape.Type;
            int refFill = SafeFill(refShape);

            var matches = new List<PPT.Shape>();
            foreach (PPT.Shape s in slide.Shapes)
                if ((int)s.Type == refType && SafeFill(s) == refFill)
                    matches.Add(s);

            if (matches.Count == 0) return "No similar shapes found.";
            bool first = true;
            foreach (var s in matches)
            {
                s.Select(first ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse);
                first = false;
            }
            return $"Selected {matches.Count} similar shape(s).";
        }

        private static int SafeFill(PPT.Shape s)
        {
            try { return s.Fill.Visible == Office.MsoTriState.msoTrue ? s.Fill.ForeColor.RGB : -1; }
            catch { return -1; }
        }

        // ---- Smart Painter (4) ----
        public static string SmartPainter()
        {
            var sr = RequireShapes(2);
            sr[1].PickUp();
            for (int i = 2; i <= sr.Count; i++) sr[i].Apply();
            return $"Painted format from the first shape onto {sr.Count - 1} other(s).";
        }

        // ---- Table of Contents (1) ----
        public static string BuildTableOfContents()
        {
            var pres = App.ActivePresentation;
            var slides = pres.Slides;
            var titles = new List<KeyValuePair<int, string>>();
            for (int i = 1; i <= slides.Count; i++)
            {
                string t = "";
                try { if (slides[i].Shapes.HasTitle == Office.MsoTriState.msoTrue) t = slides[i].Shapes.Title.TextFrame.TextRange.Text; }
                catch { }
                titles.Add(new KeyValuePair<int, string>(i, t));
            }

            var toc = TableOfContents.Build(titles, new HashSet<int> { 1 });
            if (toc.Count == 0) return "No titled slides found to build a TOC from.";

            var tocSlide = slides.Add(2, PPT.PpSlideLayout.ppLayoutTitleOnly);
            if (tocSlide.Shapes.HasTitle == Office.MsoTriState.msoTrue)
                tocSlide.Shapes.Title.TextFrame.TextRange.Text = "Table of Contents";

            float w = pres.PageSetup.SlideWidth - 80f;
            float h = pres.PageSetup.SlideHeight - 140f;
            var tb = tocSlide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 40f, 100f, w, h);
            tb.TextFrame.TextRange.Text = TableOfContents.Render(toc);
            tb.TextFrame.TextRange.Font.Name = BrandTheme.Default().Fonts.Latin;
            return $"Built a Table of Contents slide from {toc.Count} titled slides.";
        }

        // ---- Slide Check (10) ----
        public static string SlideCheck()
        {
            var pres = App.ActivePresentation;
            float sw = pres.PageSetup.SlideWidth, sh = pres.PageSetup.SlideHeight;
            var issues = new List<string>();

            for (int i = 1; i <= pres.Slides.Count; i++)
            {
                var slide = pres.Slides[i];
                bool hasTitle = false;
                try { hasTitle = slide.Shapes.HasTitle == Office.MsoTriState.msoTrue; } catch { }
                if (!hasTitle) issues.Add($"Slide {i}: no title placeholder");

                foreach (PPT.Shape s in slide.Shapes)
                {
                    var box = new LayoutBox(s.Left, s.Top, s.Width, s.Height);
                    if (SlideCheckRules.IsOffSlide(box, sw, sh))
                        issues.Add($"Slide {i}: '{s.Name}' extends off-slide");
                }
            }

            var sb = new StringBuilder();
            if (issues.Count == 0) sb.Append("Slide Check passed — no issues found.");
            else
            {
                sb.AppendLine($"Slide Check found {issues.Count} issue(s):").AppendLine();
                foreach (var line in issues.Take(25)) sb.AppendLine("• " + line);
                if (issues.Count > 25) sb.AppendLine($"…and {issues.Count - 25} more.");
            }
            MessageBox.Show(sb.ToString(), "Upslide — Slide Check", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return $"Slide Check: {issues.Count} issue(s).";
        }
    }
}
