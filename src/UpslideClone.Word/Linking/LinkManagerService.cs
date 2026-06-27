using System.Collections.Generic;
using System.IO;
using Wd = Microsoft.Office.Interop.Word;
using UpslideClone.Core.Linking;

namespace UpslideClone.Word.Linking
{
    public sealed class LinkRow
    {
        public string LinkId { get; set; }
        public string Type { get; set; }
        public string SourceFile { get; set; }
        public string SourcePath { get; set; }
        public string SheetRange { get; set; }
        public string LastRefresh { get; set; }
        public string Status { get; set; }
    }

    /// <summary>Read/repoint/navigate the Word document's link registry (FR-X3 / FR-X6).</summary>
    internal static class LinkManagerService
    {
        public static List<LinkRow> List(Wd.Document doc)
        {
            var rows = new List<LinkRow>();
            var reg = WordInterop.LoadRegistry(doc);
            foreach (var m in reg.Items)
            {
                bool srcOk = !string.IsNullOrEmpty(m.SourceWorkbook) && File.Exists(m.SourceWorkbook);
                bool anchorOk = doc.Bookmarks.Exists(LinkMetadataRegistry.BookmarkName(m.LinkId));
                rows.Add(new LinkRow
                {
                    LinkId = m.LinkId,
                    Type = m.ExportType.ToString(),
                    SourceFile = string.IsNullOrEmpty(m.SourceWorkbook) ? "(none)" : Path.GetFileName(m.SourceWorkbook),
                    SourcePath = m.SourceWorkbook,
                    SheetRange = (m.SourceSheet ?? "") + "!" + (m.SourceRange ?? ""),
                    LastRefresh = m.LastRefresh?.ToString("yyyy-MM-dd HH:mm") ?? "(never)",
                    Status = !anchorOk ? "Anchor missing" : (srcOk ? "OK" : "Source missing")
                });
            }
            return rows;
        }

        public static int ChangeSource(Wd.Document doc, ISet<string> linkIds, string newWorkbookPath)
        {
            var reg = WordInterop.LoadRegistry(doc);
            int changed = 0;
            foreach (var m in reg.Items)
                if (linkIds.Contains(m.LinkId)) { m.SourceWorkbook = newWorkbookPath; changed++; }
            if (changed > 0) WordInterop.SaveRegistry(doc, reg);
            return changed;
        }

        public static void GoTo(Wd.Document doc, string linkId)
        {
            string bm = LinkMetadataRegistry.BookmarkName(linkId);
            if (doc.Bookmarks.Exists(bm))
                doc.Bookmarks[bm].Range.Select();
        }
    }
}
