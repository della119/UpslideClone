using System;
using Office = Microsoft.Office.Core;
using PPT = Microsoft.Office.Interop.PowerPoint;
using UpslideClone.Core.Linking;

namespace UpslideClone.PowerPoint.Linking
{
    /// <summary>
    /// Sizing Guide (FR-X9): insert a tagged placeholder rectangle on the active
    /// slide. Exports that name this placeholder (Advanced Export's Placeholder
    /// column) snap to its exact position and size for consistent dimensions.
    /// </summary>
    internal static class SizingGuideCommand
    {
        public static string InsertPlaceholder(PPT.Application app)
        {
            PPT.Slide slide;
            try { slide = (PPT.Slide)app.ActiveWindow.View.Slide; }
            catch { throw new InvalidOperationException("Open a slide in Normal view first."); }

            // Next free placeholder key across the deck (P1, P2, …).
            int n = 1;
            foreach (PPT.Slide s in app.ActivePresentation.Slides)
                foreach (PPT.Shape sh in s.Shapes)
                    if (!string.IsNullOrEmpty(sh.Tags[TagKeys.Placeholder])) n++;
            string key = "P" + n;

            PPT.Shape shape = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 100, 100, 320, 200);
            shape.Tags.Add(TagKeys.Placeholder, key);
            shape.Fill.Visible = Office.MsoTriState.msoFalse;
            shape.Line.DashStyle = Office.MsoLineDashStyle.msoLineDash;
            shape.Line.ForeColor.RGB = 0x999999;
            shape.TextFrame.TextRange.Text = "Sizing placeholder " + key;
            shape.Name = "UPS_Placeholder_" + key;

            return $"Inserted sizing placeholder {key}. Use \"{key}\" in Advanced Export's Placeholder column to target it.";
        }
    }
}
