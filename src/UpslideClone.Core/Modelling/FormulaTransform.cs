using System;

namespace UpslideClone.Core.Modelling
{
    /// <summary>Pure formula text transforms (IFERROR wrapper, FR-M2).</summary>
    public static class FormulaTransform
    {
        /// <summary>
        /// Wrap a formula in IFERROR(…, replacement). Returns the input unchanged
        /// if it is not a formula or is already an IFERROR(…) at the top level.
        /// </summary>
        /// <param name="formula">e.g. "=A1/B1".</param>
        /// <param name="replacement">Literal inserted as the 2nd arg; default <c>""</c> (empty string).</param>
        public static string WrapIfError(string formula, string replacement = "\"\"")
        {
            if (string.IsNullOrWhiteSpace(formula)) return formula;

            string trimmed = formula.TrimStart();
            if (!trimmed.StartsWith("=", StringComparison.Ordinal)) return formula; // constant, not a formula

            string body = trimmed.Substring(1).Trim();
            if (body.Length == 0) return formula;

            if (body.StartsWith("IFERROR(", StringComparison.OrdinalIgnoreCase) && IsSingleTopLevelCall(body))
                return formula; // already wrapped — don't double-wrap

            return "=IFERROR(" + body + "," + replacement + ")";
        }

        /// <summary>
        /// True if <paramref name="body"/> is a single function call spanning the
        /// whole expression (its opening "(" closes only at the very end), so an
        /// existing IFERROR(...) isn't itself an argument to something else.
        /// </summary>
        private static bool IsSingleTopLevelCall(string body)
        {
            int depth = 0;
            int open = body.IndexOf('(');
            if (open < 0) return false;
            for (int i = open; i < body.Length; i++)
            {
                if (body[i] == '(') depth++;
                else if (body[i] == ')')
                {
                    depth--;
                    if (depth == 0) return i == body.Length - 1;
                }
            }
            return false;
        }
    }
}
