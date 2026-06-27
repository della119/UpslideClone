using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using PPT = Microsoft.Office.Interop.PowerPoint;
using UpslideClone.Core.Linking;

namespace UpslideClone.PowerPoint.Linking
{
    /// <summary>One linked shape discovered in a deck, with its parsed metadata.</summary>
    public sealed class LinkedShape
    {
        public int SlideIndex { get; set; }
        public PPT.Slide Slide { get; set; }
        public PPT.Shape Shape { get; set; }
        public LinkMetadata Metadata { get; set; }
    }

    /// <summary>PowerPoint-side interop helpers for the link tags + enumeration.</summary>
    internal static class PptInterop
    {
        /// <summary>Read a UPS_* tag; PowerPoint returns "" when the tag is absent.</summary>
        public static string GetTag(PPT.Shape shape, string key)
        {
            try { return shape.Tags[key] ?? ""; }
            catch { return ""; }
        }

        /// <summary>Write all of a link's tags onto a shape (Add overwrites same-named tags).</summary>
        public static void WriteTags(PPT.Shape shape, LinkMetadata meta)
        {
            foreach (var kv in meta.ToTags())
                shape.Tags.Add(kv.Key, kv.Value ?? "");
        }

        /// <summary>Enumerate every shape in the deck carrying a UPS_LinkId tag.</summary>
        public static IEnumerable<LinkedShape> EnumerateLinks(PPT.Presentation pres)
        {
            foreach (PPT.Slide slide in pres.Slides)
            {
                foreach (PPT.Shape shape in slide.Shapes)
                {
                    var meta = LinkMetadata.FromTags(k => GetTag(shape, k));
                    if (meta == null) continue;
                    yield return new LinkedShape
                    {
                        SlideIndex = slide.SlideIndex,
                        Slide = slide,
                        Shape = shape,
                        Metadata = meta
                    };
                }
            }
        }

        public static void Release(object com)
        {
            try { if (com != null && Marshal.IsComObject(com)) Marshal.ReleaseComObject(com); }
            catch { /* never throw from cleanup */ }
        }
    }

    /// <summary>
    /// Resolves source workbooks for refresh. Opens each source in a single hidden,
    /// reused Excel instance (read-only) and caches by path; closes everything on
    /// Dispose. Mirrors §6.4's "open hidden, cache &amp; reuse, close on completion".
    /// </summary>
    internal sealed class ExcelSourceResolver : IDisposable
    {
        private Excel.Application _app;
        private readonly Dictionary<string, Excel.Workbook> _open =
            new Dictionary<string, Excel.Workbook>(StringComparer.OrdinalIgnoreCase);

        private Excel.Application App()
        {
            if (_app == null)
            {
                _app = new Excel.Application { Visible = false, DisplayAlerts = false, ScreenUpdating = false };
            }
            return _app;
        }

        public bool TryGetWorkbook(string path, out Excel.Workbook wb)
        {
            wb = null;
            if (string.IsNullOrEmpty(path) || !File.Exists(path)) return false;

            if (_open.TryGetValue(path, out wb) && wb != null) return true;

            // Reuse an already-open instance if the user has the file open elsewhere.
            wb = App().Workbooks.Open(path, ReadOnly: true, UpdateLinks: false);
            _open[path] = wb;
            return true;
        }

        public Excel.Worksheet GetSheet(Excel.Workbook wb, string sheetName)
        {
            foreach (Excel.Worksheet ws in wb.Worksheets)
                if (string.Equals(ws.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                    return ws;
            return null;
        }

        public void Dispose()
        {
            foreach (var wb in _open.Values)
            {
                try { wb.Close(SaveChanges: false); } catch { }
                PptInterop.Release(wb);
            }
            _open.Clear();

            if (_app != null)
            {
                try { _app.Quit(); } catch { }
                PptInterop.Release(_app);
                _app = null;
            }
        }
    }
}
