using System;
using System.Collections.Generic;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using PPT = Microsoft.Office.Interop.PowerPoint;
using UpslideClone.Core.Linking;

namespace UpslideClone.PowerPoint.Linking
{
    /// <summary>A row shown in the Link Manager (FR-X4).</summary>
    public sealed class LinkRow
    {
        public string LinkId { get; set; }
        public int SlideIndex { get; set; }
        public string Type { get; set; }
        public string SourceFile { get; set; }   // file name only (full path in tooltip)
        public string SourcePath { get; set; }
        public string SheetRange { get; set; }
        public string LastRefresh { get; set; }
        public string Status { get; set; }
    }

    /// <summary>Read/repoint/navigate operations behind the Link Manager pane.</summary>
    internal static class LinkManagerService
    {
        public static List<LinkRow> List(PPT.Presentation pres) => List(pres, false);

        /// <summary>
        /// List all links. When <paramref name="checkDrift"/> is set, open each
        /// source and compare its current content hash to the stored one (FR-X7),
        /// reporting "Changed" when the source has moved on since last export.
        /// </summary>
        public static List<LinkRow> List(PPT.Presentation pres, bool checkDrift)
        {
            var rows = new List<LinkRow>();
            ExcelSourceResolver resolver = checkDrift ? new ExcelSourceResolver() : null;
            try
            {
                foreach (var link in PptInterop.EnumerateLinks(pres))
                {
                    var m = link.Metadata;
                    bool exists = !string.IsNullOrEmpty(m.SourceWorkbook) && File.Exists(m.SourceWorkbook);
                    string status = exists ? "OK" : "Source missing";
                    if (exists && checkDrift)
                        status = DriftStatus(resolver, m);

                    rows.Add(new LinkRow
                    {
                        LinkId = m.LinkId,
                        SlideIndex = link.SlideIndex,
                        Type = m.ExportType.ToString(),
                        SourceFile = string.IsNullOrEmpty(m.SourceWorkbook) ? "(none)" : Path.GetFileName(m.SourceWorkbook),
                        SourcePath = m.SourceWorkbook,
                        SheetRange = (m.SourceSheet ?? "") + "!" + (m.SourceRange ?? ""),
                        LastRefresh = m.LastRefresh?.ToString("yyyy-MM-dd HH:mm") ?? "(never)",
                        Status = status
                    });
                }
            }
            finally { resolver?.Dispose(); }
            return rows;
        }

        private static string DriftStatus(ExcelSourceResolver resolver, LinkMetadata m)
        {
            try
            {
                Excel.Workbook wb;
                if (!resolver.TryGetWorkbook(m.SourceWorkbook, out wb)) return "Source missing";
                Excel.Worksheet ws = resolver.GetSheet(wb, m.SourceSheet);
                if (ws == null) return "Sheet missing";

                Excel.Range range = ws.Range[m.SourceRange];
                object raw = range.Value2;
                object[,] grid = raw as object[,];
                string current = grid != null ? LinkHash.Compute(grid) : LinkHash.Compute(Convert.ToString(raw));
                return string.Equals(current, m.SourceHash, StringComparison.Ordinal) ? "OK" : "Changed";
            }
            catch { return "Error"; }
        }

        /// <summary>Repoint the given links to a new source workbook (FR-X6).</summary>
        public static int ChangeSource(PPT.Presentation pres, ISet<string> linkIds, string newWorkbookPath)
        {
            int changed = 0;
            foreach (var link in PptInterop.EnumerateLinks(pres))
            {
                if (!linkIds.Contains(link.Metadata.LinkId)) continue;
                link.Shape.Tags.Add(TagKeys.SourceWorkbook, newWorkbookPath);
                changed++;
            }
            return changed;
        }

        /// <summary>Select the slide + shape for a link in the active window.</summary>
        public static void GoTo(PPT.Application app, PPT.Presentation pres, string linkId)
        {
            foreach (var link in PptInterop.EnumerateLinks(pres))
            {
                if (link.Metadata.LinkId != linkId) continue;
                app.ActiveWindow.View.GotoSlide(link.SlideIndex);
                try { link.Shape.Select(); } catch { /* selection can fail on some views */ }
                return;
            }
        }
    }
}
