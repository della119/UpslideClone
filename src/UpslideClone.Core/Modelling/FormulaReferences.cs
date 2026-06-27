using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace UpslideClone.Core.Modelling
{
    /// <summary>A single A1 cell reference parsed from a formula.</summary>
    public sealed class CellRef
    {
        public string Raw { get; set; }
        public string Column { get; set; }   // e.g. "A", "AB"
        public int Row { get; set; }
        public bool ColumnAbsolute { get; set; }
        public bool RowAbsolute { get; set; }

        public override string ToString() => Raw;
    }

    /// <summary>
    /// Pure formula-reference engine — the basis for Fast Fill / paste-preserve
    /// (reference shifting) and Smart Track (reference extraction). No Office dep.
    ///
    /// Quoted strings ("..."), sheet names ('...') and function names are left
    /// untouched; only relative parts of A1 references shift ($-anchored parts
    /// stay put), matching Excel's relative/absolute fill semantics.
    /// </summary>
    public static class FormulaReferences
    {
        // A1 cell ref with optional $ anchors and boundary guards so we don't
        // match inside identifiers (LOG10), function calls (SUM() ) or decimals.
        private static readonly Regex CellRefRe = new Regex(
            @"(?<![A-Za-z0-9_$])(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})(?![A-Za-z0-9_(])",
            RegexOptions.Compiled);

        // Quoted string OR single-quoted sheet name — segments to skip.
        private static readonly Regex StringRe = new Regex(
            "\"(?:[^\"]|\"\")*\"|'(?:[^']|'')*'",
            RegexOptions.Compiled);

        /// <summary>Shift every relative reference in <paramref name="formula"/> by the given deltas.</summary>
        public static string Shift(string formula, int rowDelta, int colDelta)
        {
            if (string.IsNullOrEmpty(formula)) return formula;
            return ProcessCodeSegments(formula, code => CellRefRe.Replace(code, m =>
            {
                bool colAbs = m.Groups[1].Value == "$";
                string colLetters = m.Groups[2].Value;
                bool rowAbs = m.Groups[3].Value == "$";
                int row = int.Parse(m.Groups[4].Value);

                int col = A1.ColumnToIndex(colLetters);
                if (!colAbs) col = Math.Max(1, Math.Min(A1.MaxColumn, col + colDelta));
                if (!rowAbs) row = Math.Max(1, Math.Min(A1.MaxRow, row + rowDelta));

                return (colAbs ? "$" : "") + A1.IndexToColumn(col) + (rowAbs ? "$" : "") + row;
            }));
        }

        /// <summary>Extract all A1 references from a formula (for precedent tracing).</summary>
        public static IList<CellRef> Extract(string formula)
        {
            var refs = new List<CellRef>();
            if (string.IsNullOrEmpty(formula)) return refs;

            ProcessCodeSegments(formula, code =>
            {
                foreach (Match m in CellRefRe.Matches(code))
                {
                    refs.Add(new CellRef
                    {
                        Raw = m.Value,
                        ColumnAbsolute = m.Groups[1].Value == "$",
                        Column = m.Groups[2].Value.ToUpperInvariant(),
                        RowAbsolute = m.Groups[3].Value == "$",
                        Row = int.Parse(m.Groups[4].Value)
                    });
                }
                return code; // unchanged
            });

            return refs;
        }

        /// <summary>
        /// Walk the formula, applying <paramref name="transform"/> only to the
        /// segments outside quoted strings / sheet names.
        /// </summary>
        private static string ProcessCodeSegments(string input, Func<string, string> transform)
        {
            var sb = new StringBuilder();
            int pos = 0;
            foreach (Match str in StringRe.Matches(input))
            {
                if (str.Index > pos)
                    sb.Append(transform(input.Substring(pos, str.Index - pos)));
                sb.Append(str.Value);            // copy the literal verbatim
                pos = str.Index + str.Length;
            }
            if (pos < input.Length)
                sb.Append(transform(input.Substring(pos)));
            return sb.ToString();
        }
    }
}
