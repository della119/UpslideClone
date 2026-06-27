using System;
using Excel = Microsoft.Office.Interop.Excel;
using Wd = Microsoft.Office.Interop.Word;
using UpslideClone.Core.Linking;
using UpslideClone.Core.Util;

namespace UpslideClone.Word.Linking
{
    public sealed class RefreshResult
    {
        public int Ok;
        public int SourceMissing;
        public int Failed;

        public override string ToString()
        {
            return $"Refreshed {Ok} link(s)."
                 + (SourceMissing > 0 ? $" {SourceMissing} source(s)/anchor(s) missing." : "")
                 + (Failed > 0 ? $" {Failed} failed." : "");
        }
    }

    /// <summary>
    /// Re-renders Excel→Word links (FR-X3) from their sources. Each link is anchored
    /// by a bookmark UPS_&lt;id&gt;; pictures are re-pasted into the bookmark range,
    /// tables rewritten cell-by-cell. The document registry tracks source hashes.
    /// </summary>
    internal static class RefreshEngine
    {
        public static RefreshResult RefreshAll(Wd.Document doc)
        {
            var result = new RefreshResult();
            var reg = WordInterop.LoadRegistry(doc);

            using (var resolver = new ExcelSourceResolver())
            {
                foreach (var meta in reg.Items)
                {
                    try
                    {
                        string bm = LinkMetadataRegistry.BookmarkName(meta.LinkId);
                        if (!doc.Bookmarks.Exists(bm)) { result.SourceMissing++; continue; }

                        Excel.Workbook wb;
                        if (!resolver.TryGetWorkbook(meta.SourceWorkbook, out wb)) { result.SourceMissing++; continue; }
                        Excel.Worksheet ws = resolver.GetSheet(wb, meta.SourceSheet);
                        if (ws == null) { result.SourceMissing++; continue; }

                        Excel.Range range = ws.Range[meta.SourceRange];
                        object raw = range.Value2;
                        object[,] grid = raw as object[,];

                        if (meta.ExportType == ExportType.Table)
                            RefreshTable(doc, bm, range);
                        else
                            RefreshPicture(doc, bm, range);

                        meta.SourceHash = grid != null ? LinkHash.Compute(grid) : LinkHash.Compute(Convert.ToString(raw));
                        meta.LastRefresh = DateTime.Now;
                        result.Ok++;
                    }
                    catch (Exception ex)
                    {
                        Logger.Error($"Word refresh failed for link {meta.LinkId}", ex);
                        result.Failed++;
                    }
                }
            }

            WordInterop.SaveRegistry(doc, reg);
            return result;
        }

        private static void RefreshPicture(Wd.Document doc, string bm, Excel.Range range)
        {
            Wd.Range r = doc.Bookmarks[bm].Range;
            range.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlPicture);
            r.Paste();                 // replaces the bookmark range content with the fresh picture
            doc.Bookmarks.Add(bm, r);  // Paste can drop the bookmark — re-anchor it
        }

        private static void RefreshTable(Wd.Document doc, string bm, Excel.Range range)
        {
            Wd.Range r = doc.Bookmarks[bm].Range;
            if (r.Tables.Count < 1) return;
            Wd.Table tbl = r.Tables[1];

            int srcRows = range.Rows.Count, srcCols = range.Columns.Count;
            int dstRows = tbl.Rows.Count, dstCols = tbl.Columns.Count;
            for (int rr = 1; rr <= srcRows && rr <= dstRows; rr++)
                for (int c = 1; c <= srcCols && c <= dstCols; c++)
                {
                    var cell = (Excel.Range)range.Cells[rr, c];
                    string text = cell.Text == null ? "" : cell.Text.ToString();
                    tbl.Cell(rr, c).Range.Text = text;
                }
            doc.Bookmarks.Add(bm, tbl.Range);
        }
    }
}
