using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Wd = Microsoft.Office.Interop.Word;
using UpslideClone.Core.Linking;

namespace UpslideClone.Word.Linking
{
    /// <summary>
    /// Word-side helpers: the document link registry (CustomXMLPart) and a hidden
    /// Excel source resolver. Word has no Shape.Tags, so links are stored in a
    /// document-level <see cref="LinkMetadataRegistry"/> and anchored by bookmarks.
    /// </summary>
    internal static class WordInterop
    {
        private static readonly object missing = Type.Missing;

        public static LinkMetadataRegistry LoadRegistry(Wd.Document doc)
        {
            var parts = doc.CustomXMLParts.SelectByNamespace(LinkMetadataRegistry.NamespaceUri);
            if (parts.Count >= 1)
                return LinkMetadataRegistry.FromXml(parts[1].XML);
            return new LinkMetadataRegistry();
        }

        public static void SaveRegistry(Wd.Document doc, LinkMetadataRegistry reg)
        {
            var existing = doc.CustomXMLParts.SelectByNamespace(LinkMetadataRegistry.NamespaceUri);
            for (int i = existing.Count; i >= 1; i--)
                existing[i].Delete();
            doc.CustomXMLParts.Add(reg.ToXml(), missing);
        }

        public static void Release(object com)
        {
            try { if (com != null && Marshal.IsComObject(com)) Marshal.ReleaseComObject(com); }
            catch { /* never throw from cleanup */ }
        }
    }

    /// <summary>Opens source workbooks read-only in a single hidden, reused Excel instance.</summary>
    internal sealed class ExcelSourceResolver : IDisposable
    {
        private Excel.Application _app;
        private readonly Dictionary<string, Excel.Workbook> _open =
            new Dictionary<string, Excel.Workbook>(StringComparer.OrdinalIgnoreCase);

        private Excel.Application App()
        {
            if (_app == null)
                _app = new Excel.Application { Visible = false, DisplayAlerts = false, ScreenUpdating = false };
            return _app;
        }

        public bool TryGetWorkbook(string path, out Excel.Workbook wb)
        {
            wb = null;
            if (string.IsNullOrEmpty(path) || !File.Exists(path)) return false;
            if (_open.TryGetValue(path, out wb) && wb != null) return true;
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
                WordInterop.Release(wb);
            }
            _open.Clear();
            if (_app != null)
            {
                try { _app.Quit(); } catch { }
                WordInterop.Release(_app);
                _app = null;
            }
        }
    }
}
