using ExcelInterop = Microsoft.Office.Interop.Excel;
using UpslideClone.Core.Linking;

namespace UpslideClone.Excel.Commands
{
    /// <summary>
    /// Persists the <see cref="LinkedRangeRegistry"/> in the source workbook as a
    /// CustomXMLPart (§6.4), so "Highlight linked items" (FR-X10) knows which
    /// ranges feed a deck even across sessions.
    /// </summary>
    internal static class LinkRegistryStore
    {
        public static LinkedRangeRegistry Load(ExcelInterop.Workbook book)
        {
            var parts = book.CustomXMLParts.SelectByNamespace(LinkedRangeRegistry.NamespaceUri);
            if (parts.Count >= 1)
                return LinkedRangeRegistry.FromXml(parts[1].XML);
            return new LinkedRangeRegistry();
        }

        public static void Save(ExcelInterop.Workbook book, LinkedRangeRegistry registry)
        {
            // Remove any existing parts in our namespace, then add the fresh one.
            var existing = book.CustomXMLParts.SelectByNamespace(LinkedRangeRegistry.NamespaceUri);
            for (int i = existing.Count; i >= 1; i--)
                existing[i].Delete();
            book.CustomXMLParts.Add(registry.ToXml(), missing);
        }

        public static void Record(ExcelInterop.Workbook book, string linkId, string sheet, string range)
        {
            var reg = Load(book);
            reg.Upsert(linkId, sheet, range);
            Save(book, reg);
        }

        private static readonly object missing = System.Type.Missing;
    }
}
