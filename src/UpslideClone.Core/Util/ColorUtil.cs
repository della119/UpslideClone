using System;
using System.Drawing;
using System.Globalization;

namespace UpslideClone.Core.Util
{
    /// <summary>
    /// Colour helpers for Office Interop.
    /// IMPORTANT: Excel/PowerPoint Interop expect <b>BGR</b> integers, not RGB.
    /// Always go through <see cref="ToOle"/> (== ColorTranslator.ToOle) before
    /// assigning to <c>.Color</c> / <c>.RGB</c> Interop properties.
    /// </summary>
    public static class ColorUtil
    {
        /// <summary>Parse "#RRGGBB" (or "RRGGBB") into a <see cref="Color"/>.</summary>
        public static Color FromHex(string hex)
        {
            if (string.IsNullOrWhiteSpace(hex))
                throw new ArgumentException("Empty colour string.", nameof(hex));

            var s = hex.Trim().TrimStart('#');
            if (s.Length != 6)
                throw new FormatException($"Expected a 6-digit hex colour, got '{hex}'.");

            int r = int.Parse(s.Substring(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            int g = int.Parse(s.Substring(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            int b = int.Parse(s.Substring(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            return Color.FromArgb(r, g, b);
        }

        /// <summary>BGR integer for Office Interop (== System.Drawing.ColorTranslator.ToOle).</summary>
        public static int ToOle(Color c)
        {
            return ColorTranslator.ToOle(c);
        }

        /// <summary>Convenience: "#RRGGBB" → BGR OLE integer for Interop.</summary>
        public static int OleFromHex(string hex)
        {
            return ToOle(FromHex(hex));
        }
    }
}
