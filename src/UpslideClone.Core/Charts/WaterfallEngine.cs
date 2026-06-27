using System;
using System.Collections.Generic;

namespace UpslideClone.Core.Charts
{
    /// <summary>A single label/value pair read from the source bridge table.</summary>
    public sealed class WaterfallPoint
    {
        public string Label { get; set; }
        public double Value { get; set; }

        public WaterfallPoint() { }

        public WaterfallPoint(string label, double value)
        {
            Label = label;
            Value = value;
        }
    }

    /// <summary>
    /// One computed row of the stacked-column helper block.
    /// A null Decrease/Increase/Total means "#N/A" (=NA()) — Excel draws no bar
    /// and no zero-label for that series, exactly like the UPSLIDE_Waterfall tab.
    /// </summary>
    public sealed class WaterfallRow
    {
        public string Label { get; set; }
        public double Base { get; set; }
        public double? Decrease { get; set; }
        public double? Increase { get; set; }
        public double? Total { get; set; }
    }

    /// <summary>
    /// Pure waterfall geometry — no Office dependency, fully unit-testable.
    /// Faithful port of addin/src/core/waterfall.ts:computeWaterfall.
    ///
    /// Reproduces Upslide's waterfall WITHOUT a native waterfall chart type:
    /// a standard stacked column chart where an invisible "Base" series floats
    /// each delta bar to the correct height.
    ///
    /// Anchor (total) rows are auto-detected: a row whose value equals the
    /// running cumulative of the preceding deltas is treated as a total
    /// (full-height bar). The first row is always an anchor.
    /// </summary>
    public static class WaterfallEngine
    {
        private const double AnchorTolerance = 1e-6;

        public static List<WaterfallRow> Compute(IList<WaterfallPoint> points)
        {
            if (points == null) throw new ArgumentNullException(nameof(points));

            var rows = new List<WaterfallRow>(points.Count);
            double running = 0;

            for (int i = 0; i < points.Count; i++)
            {
                var p = points[i];
                bool isAnchor = i == 0 || Math.Abs(p.Value - running) < AnchorTolerance;

                if (isAnchor)
                {
                    running = p.Value;
                    rows.Add(new WaterfallRow
                    {
                        Label = p.Label,
                        Base = 0,
                        Decrease = null,
                        Increase = null,
                        Total = p.Value
                    });
                }
                else if (p.Value >= 0)
                {
                    rows.Add(new WaterfallRow
                    {
                        Label = p.Label,
                        Base = running,
                        Decrease = null,
                        Increase = p.Value,
                        Total = null
                    });
                    running += p.Value;
                }
                else
                {
                    running += p.Value; // running now at the new (lower) level
                    rows.Add(new WaterfallRow
                    {
                        Label = p.Label,
                        Base = running,
                        Decrease = -p.Value,
                        Increase = null,
                        Total = null
                    });
                }
            }

            return rows;
        }
    }
}
