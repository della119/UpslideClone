using System;
using Office = Microsoft.Office.Core;
using PPT = Microsoft.Office.Interop.PowerPoint;
using UpslideClone.Core.Design;

namespace UpslideClone.PowerPoint.Design
{
    /// <summary>
    /// Cross-references &amp; footnotes (#2). Cross-references store the target's
    /// stable SlideID on the shape so they survive re-ordering and re-render on
    /// refresh; footnotes are auto-numbered per slide.
    /// </summary>
    internal static class ReferenceCommands
    {
        private static readonly Office.MsoTextOrientation H = Office.MsoTextOrientation.msoTextOrientationHorizontal;

        public static string InsertFootnote(PPT.Application app)
        {
            var pres = app.ActivePresentation;
            var slide = (PPT.Slide)app.ActiveWindow.View.Slide;

            int count = 0;
            foreach (PPT.Shape s in slide.Shapes)
                if (!string.IsNullOrEmpty(s.Tags[CrossReference.FootnoteTag])) count++;
            int number = count + 1;

            string text = Prompt.Input("Insert Footnote", "Footnote text:");
            if (text == null) return "Cancelled.";

            float w = pres.PageSetup.SlideWidth - 80f;
            float top = pres.PageSetup.SlideHeight - 36f;
            var tb = slide.Shapes.AddTextbox(H, 40f, top, w, 24f);
            tb.TextFrame.TextRange.Text = CrossReference.Footnote(number, text);
            tb.TextFrame.TextRange.Font.Size = 9f;
            tb.Tags.Add(CrossReference.FootnoteTag, number.ToString());
            return $"Inserted footnote {number}.";
        }

        public static string InsertCrossReference(PPT.Application app)
        {
            var pres = app.ActivePresentation;
            string s = Prompt.Input("Insert Cross-reference", $"Target slide number (1-{pres.Slides.Count}):");
            if (s == null) return "Cancelled.";
            int n;
            if (!int.TryParse(s.Trim(), out n) || n < 1 || n > pres.Slides.Count)
                throw new InvalidOperationException("Enter a valid slide number.");

            var target = pres.Slides[n];
            var slide = (PPT.Slide)app.ActiveWindow.View.Slide;
            var tb = slide.Shapes.AddTextbox(H, 40f, 40f, 320f, 24f);
            tb.TextFrame.TextRange.Text = CrossReference.Format(n, TitleOf(target));
            tb.Tags.Add(CrossReference.TargetIdTag, target.SlideID.ToString());
            return "Inserted cross-reference.";
        }

        public static string RefreshCrossReferences(PPT.Application app)
        {
            var pres = app.ActivePresentation;
            int updated = 0;
            foreach (PPT.Slide slide in pres.Slides)
                foreach (PPT.Shape shape in slide.Shapes)
                {
                    string idTag = shape.Tags[CrossReference.TargetIdTag];
                    if (string.IsNullOrEmpty(idTag)) continue;
                    int id;
                    if (!int.TryParse(idTag, out id)) continue;
                    var target = FindBySlideId(pres, id);
                    if (target == null) continue;
                    try { shape.TextFrame.TextRange.Text = CrossReference.Format(target.SlideIndex, TitleOf(target)); updated++; }
                    catch { }
                }
            return $"Refreshed {updated} cross-reference(s).";
        }

        private static string TitleOf(PPT.Slide s)
        {
            try { if (s.Shapes.HasTitle == Office.MsoTriState.msoTrue) return s.Shapes.Title.TextFrame.TextRange.Text; }
            catch { }
            return "";
        }

        private static PPT.Slide FindBySlideId(PPT.Presentation p, int id)
        {
            foreach (PPT.Slide s in p.Slides) if (s.SlideID == id) return s;
            return null;
        }
    }
}
