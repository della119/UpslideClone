using System;
using System.Runtime.InteropServices;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using PPT = Microsoft.Office.Interop.PowerPoint;
using UpslideClone.Core.Linking;

namespace UpslideClone.Excel.Commands
{
    /// <summary>
    /// Export a range to PowerPoint as a LINKED object (FR-X1 picture, FR-X2 table),
    /// tagged with UPS_* metadata, recorded in the workbook's link registry (FR-X10),
    /// and optionally snapped to a Sizing-Guide placeholder (FR-X9).
    /// </summary>
    internal static class ExportToPowerPointCommand
    {
        private static readonly object missing = Type.Missing;

        public static string Run(ExportType type)
        {
            var sel = ExcelHelpers.RequireRangeSelection();
            var app = GetOrStartPowerPoint();
            PPT.Presentation pres = ActivePresentation(app);
            PPT.Shape shape = ExportRangeToSlide(app, pres, sel, type, 0, null);
            try { app.Activate(); } catch { }
            return $"Exported a linked {type} to slide {shape.SlideIndexOf()}.";
        }

        /// <summary>Core export used by single export and Advanced Export.</summary>
        public static PPT.Shape ExportRangeToSlide(
            PPT.Application app, PPT.Presentation pres, ExcelInterop.Range range,
            ExportType type, int targetSlideIndex, string placeholderKey)
        {
            var sheet = (ExcelInterop.Worksheet)range.Worksheet;
            var book = (ExcelInterop.Workbook)sheet.Parent;
            if (string.IsNullOrEmpty(book.Path))
                throw new InvalidOperationException("Save the workbook first so links have a source file to refresh from.");

            PPT.Slide slide = ResolveSlide(app, pres, targetSlideIndex);

            object raw = range.Value2;
            object[,] grid = raw as object[,];
            var meta = new LinkMetadata
            {
                LinkId = LinkMetadata.NewId(),
                SourceWorkbook = book.FullName,
                SourceSheet = sheet.Name,
                SourceRange = range.get_Address(true, true, ExcelInterop.XlReferenceStyle.xlA1, missing, missing),
                ExportType = type,
                SourceHash = grid != null ? LinkHash.Compute(grid) : LinkHash.Compute(Convert.ToString(raw)),
                LastRefresh = DateTime.Now
            };

            PPT.Shape shape = type == ExportType.Table ? CreateTable(slide, range) : CreatePicture(slide, range);
            SnapToPlaceholder(slide, shape, placeholderKey);

            foreach (var kv in meta.ToTags())
                shape.Tags.Add(kv.Key, kv.Value ?? "");

            // Mirror the linked range into the workbook registry (FR-X10).
            try { LinkRegistryStore.Record(book, meta.LinkId, meta.SourceSheet, meta.SourceRange); }
            catch { /* registry is best-effort; never block the export */ }

            return shape;
        }

        private static PPT.Slide ResolveSlide(PPT.Application app, PPT.Presentation pres, int targetSlideIndex)
        {
            if (targetSlideIndex > 0)
            {
                while (pres.Slides.Count < targetSlideIndex)
                    pres.Slides.Add(pres.Slides.Count + 1, PPT.PpSlideLayout.ppLayoutBlank);
                return pres.Slides[targetSlideIndex];
            }
            return ActiveOrNewSlide(app, pres);
        }

        private static void SnapToPlaceholder(PPT.Slide slide, PPT.Shape shape, string placeholderKey)
        {
            if (string.IsNullOrEmpty(placeholderKey)) return;
            foreach (PPT.Shape s in slide.Shapes)
            {
                if (string.Equals(s.Tags[TagKeys.Placeholder], placeholderKey, StringComparison.OrdinalIgnoreCase))
                {
                    shape.LockAspectRatio = Office.MsoTriState.msoFalse;
                    shape.Left = s.Left; shape.Top = s.Top; shape.Width = s.Width; shape.Height = s.Height;
                    return;
                }
            }
        }

        internal static PPT.Application GetOrStartPowerPoint()
        {
            try { return (PPT.Application)Marshal.GetActiveObject("PowerPoint.Application"); }
            catch (COMException) { return new PPT.Application(); }
        }

        internal static PPT.Presentation ActivePresentation(PPT.Application app)
        {
            return app.Presentations.Count > 0 ? app.ActivePresentation : app.Presentations.Add(Office.MsoTriState.msoTrue);
        }

        private static PPT.Slide ActiveOrNewSlide(PPT.Application app, PPT.Presentation pres)
        {
            try
            {
                var view = app.ActiveWindow.View;
                if (view?.Slide is PPT.Slide s) return s;
            }
            catch { /* no active window yet */ }

            if (pres.Slides.Count == 0)
                return pres.Slides.Add(1, PPT.PpSlideLayout.ppLayoutBlank);
            return pres.Slides[pres.Slides.Count];
        }

        private static PPT.Shape CreatePicture(PPT.Slide slide, ExcelInterop.Range sel)
        {
            sel.CopyPicture(ExcelInterop.XlPictureAppearance.xlScreen, ExcelInterop.XlCopyPictureFormat.xlPicture);
            PPT.ShapeRange pasted = slide.Shapes.Paste();
            PPT.Shape shape = pasted[1];
            shape.Left = 40; shape.Top = 80;
            return shape;
        }

        private static PPT.Shape CreateTable(PPT.Slide slide, ExcelInterop.Range sel)
        {
            int rows = sel.Rows.Count, cols = sel.Columns.Count;
            PPT.Shape shape = slide.Shapes.AddTable(rows, cols, 40, 80, 600, 30 * rows);
            PPT.Table tbl = shape.Table;
            for (int r = 1; r <= rows; r++)
                for (int c = 1; c <= cols; c++)
                {
                    var cell = (ExcelInterop.Range)sel.Cells[r, c];
                    string text = cell.Text == null ? "" : cell.Text.ToString();
                    tbl.Cell(r, c).Shape.TextFrame.TextRange.Text = text;
                }
            return shape;
        }
    }

    internal static class ShapeExtensions
    {
        /// <summary>The slide index a shape sits on (via its parent slide).</summary>
        public static int SlideIndexOf(this PPT.Shape shape) => ((PPT.Slide)shape.Parent).SlideIndex;
    }
}
