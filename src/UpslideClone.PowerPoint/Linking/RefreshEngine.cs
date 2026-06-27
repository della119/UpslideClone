using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using PPT = Microsoft.Office.Interop.PowerPoint;
using UpslideClone.Core.Linking;
using UpslideClone.Core.Util;

namespace UpslideClone.PowerPoint.Linking
{
    public sealed class RefreshResult
    {
        public int Ok;
        public int SourceMissing;
        public int Failed;

        public override string ToString()
        {
            return $"Refreshed {Ok} link(s)."
                 + (SourceMissing > 0 ? $" {SourceMissing} source(s) missing." : "")
                 + (Failed > 0 ? $" {Failed} failed." : "");
        }
    }

    /// <summary>
    /// Re-renders linked objects from their Excel sources (FR-X5), preserving each
    /// shape's position and size. Tables are rewritten in place; pictures are
    /// re-pasted and snapped back to the old geometry.
    /// </summary>
    internal static class RefreshEngine
    {
        public static RefreshResult RefreshAll(PPT.Presentation pres) => Run(pres, null);

        public static RefreshResult RefreshLinks(PPT.Presentation pres, ISet<string> linkIds) => Run(pres, linkIds);

        private static RefreshResult Run(PPT.Presentation pres, ISet<string> filter)
        {
            var result = new RefreshResult();
            using (var resolver = new ExcelSourceResolver())
            {
                // Snapshot first — refreshing a picture replaces the shape mid-enumeration.
                var links = new List<LinkedShape>(PptInterop.EnumerateLinks(pres));
                foreach (var link in links)
                {
                    if (filter != null && !filter.Contains(link.Metadata.LinkId)) continue;
                    try
                    {
                        var status = RefreshOne(link, resolver);
                        if (status == RefreshStatus.Ok) result.Ok++;
                        else if (status == RefreshStatus.SourceMissing) result.SourceMissing++;
                        else result.Failed++;
                    }
                    catch (Exception ex)
                    {
                        Logger.Error($"Refresh failed for link {link.Metadata.LinkId}", ex);
                        result.Failed++;
                    }
                }
            }
            return result;
        }

        private enum RefreshStatus { Ok, SourceMissing, Failed }

        private static RefreshStatus RefreshOne(LinkedShape link, ExcelSourceResolver resolver)
        {
            var meta = link.Metadata;
            Excel.Workbook wb;
            if (!resolver.TryGetWorkbook(meta.SourceWorkbook, out wb))
                return RefreshStatus.SourceMissing;

            Excel.Worksheet ws = resolver.GetSheet(wb, meta.SourceSheet);
            if (ws == null) return RefreshStatus.SourceMissing;

            Excel.Range range = ws.Range[meta.SourceRange];

            // New content hash for drift tracking.
            object raw = range.Value2;
            object[,] grid = raw as object[,];
            string newHash = grid != null ? LinkHash.Compute(grid) : LinkHash.Compute(Convert.ToString(raw));

            bool isTable = link.Shape.HasTable == Office.MsoTriState.msoTrue
                           && meta.ExportType == ExportType.Table;

            if (isTable)
                RefreshTable(link.Shape, range);
            else
                link.Shape = RefreshPicture(link.Slide, link.Shape, range, meta);

            meta.LastRefresh = DateTime.Now;
            meta.SourceHash = newHash;
            PptInterop.WriteTags(link.Shape, meta);
            return RefreshStatus.Ok;
        }

        private static void RefreshTable(PPT.Shape shape, Excel.Range range)
        {
            PPT.Table tbl = shape.Table;
            int srcRows = range.Rows.Count, srcCols = range.Columns.Count;
            int dstRows = tbl.Rows.Count, dstCols = tbl.Columns.Count;

            for (int r = 1; r <= srcRows && r <= dstRows; r++)
                for (int c = 1; c <= srcCols && c <= dstCols; c++)
                {
                    var cell = (Excel.Range)range.Cells[r, c];
                    string text = cell.Text == null ? "" : cell.Text.ToString();
                    tbl.Cell(r, c).Shape.TextFrame.TextRange.Text = text;
                }
        }

        private static PPT.Shape RefreshPicture(PPT.Slide slide, PPT.Shape oldShape, Excel.Range range, LinkMetadata meta)
        {
            float left = oldShape.Left, top = oldShape.Top, width = oldShape.Width, height = oldShape.Height;

            range.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlPicture);
            PPT.ShapeRange pasted = slide.Shapes.Paste();
            PPT.Shape newShape = pasted[1];

            newShape.LockAspectRatio = Office.MsoTriState.msoFalse;
            newShape.Left = left;
            newShape.Top = top;
            newShape.Width = width;
            newShape.Height = height;

            oldShape.Delete();
            return newShape;
        }
    }
}
