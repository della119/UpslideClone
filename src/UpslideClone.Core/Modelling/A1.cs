using System;
using System.Text;

namespace UpslideClone.Core.Modelling
{
    /// <summary>
    /// A1-notation helpers (column letters ↔ index). Pure, unit-testable.
    /// Used by the Fast Fill / paste-preserve reference shifter.
    /// </summary>
    public static class A1
    {
        /// <summary>"A" → 1, "Z" → 26, "AA" → 27, "XFD" → 16384.</summary>
        public static int ColumnToIndex(string letters)
        {
            if (string.IsNullOrEmpty(letters)) throw new ArgumentException("Empty column.", nameof(letters));
            int n = 0;
            foreach (char ch in letters)
            {
                char c = char.ToUpperInvariant(ch);
                if (c < 'A' || c > 'Z') throw new FormatException("Invalid column letter: " + letters);
                n = n * 26 + (c - 'A' + 1);
            }
            return n;
        }

        /// <summary>1 → "A", 27 → "AA". Clamps to the Excel max column (16384 / XFD).</summary>
        public static string IndexToColumn(int index)
        {
            if (index < 1) throw new ArgumentOutOfRangeException(nameof(index), "Column index is 1-based.");
            if (index > 16384) index = 16384;
            var sb = new StringBuilder();
            while (index > 0)
            {
                int rem = (index - 1) % 26;
                sb.Insert(0, (char)('A' + rem));
                index = (index - 1) / 26;
            }
            return sb.ToString();
        }

        public const int MaxRow = 1048576;
        public const int MaxColumn = 16384;
    }
}
