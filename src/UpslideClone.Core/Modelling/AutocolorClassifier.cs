using System.Text.RegularExpressions;

namespace UpslideClone.Core.Modelling
{
    /// <summary>Banker colour-convention class for a cell (FR-M1).</summary>
    public enum CellColorClass
    {
        Empty,
        Input,   // hardcoded constant → blue
        Formula, // same-sheet formula → black
        Link     // cross-sheet / external reference → green
    }

    /// <summary>
    /// Autocolor classifier (FR-M1) — pure. Blue = hardcoded inputs,
    /// black = formulas, green = links/external refs (cross-sheet or other-workbook).
    /// Colour application is done by the Excel command via the theme; this only
    /// decides the class.
    /// </summary>
    public static class AutocolorClassifier
    {
        // A sheet/workbook qualifier shows up as a "!" separator (e.g. Sheet1!A1,
        // [Book.xlsx]Sheet1!A1). External workbooks add [...]; both imply a link.
        private static readonly Regex LinkRe = new Regex(@"!|\[[^\]]+\]", RegexOptions.Compiled);

        /// <summary>Classic banker hex per class (blue / black / green).</summary>
        public static string DefaultHex(CellColorClass cls)
        {
            switch (cls)
            {
                case CellColorClass.Input: return "#0000FF";
                case CellColorClass.Link: return "#008000";
                case CellColorClass.Formula: return "#000000";
                default: return "#000000";
            }
        }

        /// <param name="isFormula">True if the cell holds a formula (Excel: starts with '=').</param>
        /// <param name="formula">The formula text (with or without leading '='); ignored when not a formula.</param>
        /// <param name="isEmpty">True if the cell has no value at all.</param>
        public static CellColorClass Classify(bool isFormula, string formula, bool isEmpty = false)
        {
            if (isEmpty) return CellColorClass.Empty;
            if (!isFormula) return CellColorClass.Input;
            if (!string.IsNullOrEmpty(formula) && LinkRe.IsMatch(formula)) return CellColorClass.Link;
            return CellColorClass.Formula;
        }
    }
}
