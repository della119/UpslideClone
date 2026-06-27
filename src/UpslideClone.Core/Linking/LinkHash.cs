using System;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;

namespace UpslideClone.Core.Linking
{
    /// <summary>
    /// Stable content hash of a source range's values (FR-X7 versioning).
    /// Deterministic and culture-invariant so the same data always hashes the
    /// same; the Link Manager compares the stored hash to the live one to flag
    /// "source changed since last export". Pure — no Office dependency.
    /// </summary>
    public static class LinkHash
    {
        /// <summary>Hash a 2-D value block (row-major). Null cells are encoded distinctly from empty strings.</summary>
        public static string Compute(object[,] values)
        {
            var sb = new StringBuilder();
            if (values != null)
            {
                int r0 = values.GetLowerBound(0), r1 = values.GetUpperBound(0);
                int c0 = values.GetLowerBound(1), c1 = values.GetUpperBound(1);
                for (int r = r0; r <= r1; r++)
                {
                    for (int c = c0; c <= c1; c++)
                        AppendCell(sb, values[r, c]);
                    sb.Append('␞'); // record separator between rows
                }
            }
            return Hash(sb.ToString());
        }

        /// <summary>Hash a pre-rendered string (e.g. for text exports).</summary>
        public static string Compute(string content)
        {
            return Hash(content ?? "");
        }

        private static void AppendCell(StringBuilder sb, object v)
        {
            if (v == null) sb.Append("␀");          // NUL marker = empty cell
            else if (v is double d) sb.Append(d.ToString("R", CultureInfo.InvariantCulture));
            else if (v is bool b) sb.Append(b ? "1" : "0");
            else sb.Append(Convert.ToString(v, CultureInfo.InvariantCulture));
            sb.Append('␟'); // unit separator between cells
        }

        private static string Hash(string s)
        {
            using (var sha = SHA256.Create())
            {
                byte[] bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(s));
                var hex = new StringBuilder(bytes.Length * 2);
                foreach (byte bt in bytes) hex.Append(bt.ToString("x2"));
                return hex.ToString();
            }
        }
    }
}
