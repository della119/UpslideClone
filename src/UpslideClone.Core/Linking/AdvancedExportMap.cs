using System;
using System.Collections.Generic;
using System.Globalization;

namespace UpslideClone.Core.Linking
{
    /// <summary>One row of an Advanced Export mapping: a source range → a target slide.</summary>
    public sealed class ExportMapEntry
    {
        public string SourceSheet { get; set; }
        public string SourceRange { get; set; }
        public int TargetSlide { get; set; }      // 1-based; 0 = append to end
        public ExportType ExportType { get; set; }
        public string PlaceholderKey { get; set; } // optional Sizing-Guide target
    }

    /// <summary>
    /// Parses the Advanced Export mapping table (FR-X8) — a header row naming the
    /// columns plus one row per range→slide mapping. Column order is free; columns
    /// are matched by header keyword (sheet / range / slide / type / placeholder).
    /// Pure and unit-testable.
    /// </summary>
    public static class AdvancedExportMap
    {
        public static List<ExportMapEntry> Parse(object[,] grid)
        {
            var entries = new List<ExportMapEntry>();
            if (grid == null) return entries;

            int r0 = grid.GetLowerBound(0), r1 = grid.GetUpperBound(0);
            int c0 = grid.GetLowerBound(1), c1 = grid.GetUpperBound(1);
            if (r1 - r0 < 1) return entries; // need header + ≥1 row

            // Map columns by header keyword.
            int sheetCol = -1, rangeCol = -1, slideCol = -1, typeCol = -1, phCol = -1;
            for (int c = c0; c <= c1; c++)
            {
                string h = (grid[r0, c]?.ToString() ?? "").ToLowerInvariant();
                if (h.Contains("sheet")) sheetCol = c;
                else if (h.Contains("range")) rangeCol = c;
                else if (h.Contains("slide")) slideCol = c;
                else if (h.Contains("type")) typeCol = c;
                else if (h.Contains("placeholder") || h.Contains("size")) phCol = c;
            }
            if (rangeCol < 0)
                throw new FormatException("Advanced Export map needs a 'Range' column in the header row.");

            for (int r = r0 + 1; r <= r1; r++)
            {
                string range = Cell(grid, r, rangeCol);
                if (string.IsNullOrWhiteSpace(range)) continue; // skip blank rows

                entries.Add(new ExportMapEntry
                {
                    SourceSheet = sheetCol >= 0 ? NullIfEmpty(Cell(grid, r, sheetCol)) : null,
                    SourceRange = range.Trim(),
                    TargetSlide = ParseSlide(slideCol >= 0 ? Cell(grid, r, slideCol) : ""),
                    ExportType = ParseType(typeCol >= 0 ? Cell(grid, r, typeCol) : ""),
                    PlaceholderKey = phCol >= 0 ? NullIfEmpty(Cell(grid, r, phCol)) : null
                });
            }
            return entries;
        }

        private static string Cell(object[,] g, int r, int c) => g[r, c]?.ToString() ?? "";

        private static int ParseSlide(string s)
        {
            int v;
            return int.TryParse((s ?? "").Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out v) && v > 0 ? v : 0;
        }

        private static ExportType ParseType(string s)
        {
            s = (s ?? "").Trim().ToLowerInvariant();
            return s.StartsWith("tab") ? ExportType.Table : ExportType.Picture;
        }

        private static string NullIfEmpty(string s) => string.IsNullOrWhiteSpace(s) ? null : s.Trim();
    }
}
