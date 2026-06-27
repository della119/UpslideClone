using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace UpslideClone.Core.Formatting
{
    /// <summary>
    /// Pure Smart Format heuristics — no Office dependency, fully unit-testable.
    /// Faithful port of addin/src/core/smartFormat.ts (RESULT_RE, isPercentColumn).
    /// The result-row keyword set is overridable from theme.json (EN + CN).
    /// </summary>
    public static class SmartFormatRules
    {
        /// <summary>
        /// Default result/total-row detector (EN + CN). Mirrors the prototype's RESULT_RE
        /// plus the Chinese keywords from the spec (毛利 / 净利润 / 营业利润 / 合计).
        /// </summary>
        public static readonly Regex DefaultResultRow = new Regex(
            @"\b(ebitda|ebit|gross margin|net income|net profit|operating (income|profit)|total|net sales|grand total)\b"
            + @"|毛利|净利润|营业利润|合计",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        /// <summary>True if <paramref name="label"/> names a result/total row using the default detector.</summary>
        public static bool IsResultRow(string label)
        {
            return !string.IsNullOrEmpty(label) && DefaultResultRow.IsMatch(label);
        }

        /// <summary>
        /// True if <paramref name="label"/> matches any caller-supplied keyword (case-insensitive),
        /// used when keywords are loaded from theme.json. Falls back to the default detector
        /// when <paramref name="keywords"/> is null/empty.
        /// </summary>
        public static bool IsResultRow(string label, IEnumerable<string> keywords)
        {
            if (string.IsNullOrEmpty(label)) return false;
            if (keywords == null) return IsResultRow(label);

            var list = keywords.Where(k => !string.IsNullOrWhiteSpace(k)).ToList();
            if (list.Count == 0) return IsResultRow(label);

            foreach (var k in list)
                if (label.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
            return false;
        }

        /// <summary>
        /// Decide whether a data column holds ratios (→ percent format).
        /// Port of isPercentColumn: header keyword match, OR every value is a
        /// non-integer with magnitude &lt;= 1.
        /// </summary>
        public static bool IsPercentColumn(string headerText, IList<double> values)
        {
            if (!string.IsNullOrEmpty(headerText)
                && Regex.IsMatch(headerText, @"%|cagr|margin|growth|rate|ratio", RegexOptions.IgnoreCase))
                return true;

            var nums = (values ?? new List<double>()).Where(v => !double.IsNaN(v)).ToList();
            if (nums.Count == 0) return false;

            return nums.All(v => Math.Abs(v) <= 1 && v != Math.Truncate(v));
        }
    }
}
