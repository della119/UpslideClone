using System;
using System.Runtime.InteropServices;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using UpslideClone.Core.Linking;

namespace UpslideClone.Excel.Commands
{
    /// <summary>
    /// Export the selected range to Word as a LINKED object (FR-X3). The object is
    /// inserted at Word's cursor, anchored by a bookmark UPS_&lt;id&gt;, and its
    /// metadata stored in the document's CustomXMLPart registry so the Word add-in
    /// can refresh it.
    /// </summary>
    internal static class ExportToWordCommand
    {
        private static readonly object missing = Type.Missing;

        public static string Run(ExportType type)
        {
            var sel = ExcelHelpers.RequireRangeSelection();
            var sheet = (ExcelInterop.Worksheet)sel.Worksheet;
            var book = (ExcelInterop.Workbook)sheet.Parent;
            if (string.IsNullOrEmpty(book.Path))
                throw new InvalidOperationException("Save the workbook first so the link has a source file to refresh from.");

            Word.Application app = GetOrStartWord();
            Word.Document doc = app.Documents.Count > 0 ? app.ActiveDocument : app.Documents.Add(ref missingRef, ref missingRef, ref missingRef, ref missingRef);

            object raw = sel.Value2;
            object[,] grid = raw as object[,];
            var meta = new LinkMetadata
            {
                LinkId = LinkMetadata.NewId(),
                SourceWorkbook = book.FullName,
                SourceSheet = sheet.Name,
                SourceRange = sel.get_Address(true, true, ExcelInterop.XlReferenceStyle.xlA1, missing, missing),
                ExportType = type,
                SourceHash = grid != null ? LinkHash.Compute(grid) : LinkHash.Compute(Convert.ToString(raw)),
                LastRefresh = DateTime.Now
            };
            string bm = LinkMetadataRegistry.BookmarkName(meta.LinkId);

            Word.Range insertion = app.Selection.Range;
            insertion.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

            if (type == ExportType.Table)
                InsertTable(doc, insertion, sel, bm);
            else
                InsertPicture(doc, insertion, sel, bm);

            // Store metadata in the document registry.
            var reg = LoadRegistry(doc);
            reg.Upsert(meta);
            SaveRegistry(doc, reg);

            try { app.Activate(); } catch { }
            return $"Exported a linked {type} to Word.";
        }

        private static object missingRef = Type.Missing;

        private static Word.Application GetOrStartWord()
        {
            try { return (Word.Application)Marshal.GetActiveObject("Word.Application"); }
            catch (COMException) { return new Word.Application { Visible = true }; }
        }

        private static void InsertPicture(Word.Document doc, Word.Range insertion, ExcelInterop.Range sel, string bm)
        {
            sel.CopyPicture(ExcelInterop.XlPictureAppearance.xlScreen, ExcelInterop.XlCopyPictureFormat.xlPicture);
            insertion.Paste();                 // insertion now spans the pasted picture
            doc.Bookmarks.Add(bm, insertion);
        }

        private static void InsertTable(Word.Document doc, Word.Range insertion, ExcelInterop.Range sel, string bm)
        {
            int rows = sel.Rows.Count, cols = sel.Columns.Count;
            Word.Table tbl = doc.Tables.Add(insertion, rows, cols, ref missingRef, ref missingRef);
            for (int r = 1; r <= rows; r++)
                for (int c = 1; c <= cols; c++)
                {
                    var cell = (ExcelInterop.Range)sel.Cells[r, c];
                    tbl.Cell(r, c).Range.Text = cell.Text == null ? "" : cell.Text.ToString();
                }
            doc.Bookmarks.Add(bm, tbl.Range);
        }

        // --- document registry (Word CustomXMLPart) ---

        private static LinkMetadataRegistry LoadRegistry(Word.Document doc)
        {
            var parts = doc.CustomXMLParts.SelectByNamespace(LinkMetadataRegistry.NamespaceUri);
            if (parts.Count >= 1)
                return LinkMetadataRegistry.FromXml(parts[1].XML);
            return new LinkMetadataRegistry();
        }

        private static void SaveRegistry(Word.Document doc, LinkMetadataRegistry reg)
        {
            var existing = doc.CustomXMLParts.SelectByNamespace(LinkMetadataRegistry.NamespaceUri);
            for (int i = existing.Count; i >= 1; i--)
                existing[i].Delete();
            doc.CustomXMLParts.Add(reg.ToXml(), missing);
        }
    }
}
