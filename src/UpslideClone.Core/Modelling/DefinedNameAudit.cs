using System.Text.RegularExpressions;

namespace UpslideClone.Core.Modelling
{
    /// <summary>
    /// Pure helpers for Clean &amp; Prepare (FR-M6): classify a workbook defined
    /// name by its RefersTo formula so the Excel command can prune broken ones.
    /// </summary>
    public static class DefinedNameAudit
    {
        /// <summary>A name is broken if its RefersTo points at a deleted reference (#REF!).</summary>
        public static bool IsBroken(string refersTo)
        {
            return !string.IsNullOrEmpty(refersTo) && refersTo.IndexOf("#REF!", System.StringComparison.OrdinalIgnoreCase) >= 0;
        }

        /// <summary>External/link names point at another workbook ("[Book.xlsx]" or a path).</summary>
        public static bool IsExternal(string refersTo)
        {
            if (string.IsNullOrEmpty(refersTo)) return false;
            return Regex.IsMatch(refersTo, @"\[[^\]]+\]") || refersTo.Contains(":\\") || refersTo.Contains("//");
        }
    }
}
