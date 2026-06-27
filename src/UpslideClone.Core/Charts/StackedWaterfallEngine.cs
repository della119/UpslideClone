using System;
using System.Collections.Generic;
using System.Linq;

namespace UpslideClone.Core.Charts
{
    /// <summary>A bridge step split across categories (e.g. FR / UK / DE).</summary>
    public sealed class StackedWaterfallPoint
    {
        public string Label { get; set; }
        /// <summary>One value per category, aligned to <see cref="StackedWaterfallResult.Categories"/>.</summary>
        public double[] Values { get; set; }

        public StackedWaterfallPoint() { }

        public StackedWaterfallPoint(string label, double[] values)
        {
            Label = label;
            Values = values;
        }
    }

    /// <summary>One computed row: an invisible Base float plus a segment per category.</summary>
    public sealed class StackedWaterfallRow
    {
        public string Label { get; set; }
        public double Base { get; set; }
        /// <summary>Visible segment per category; null = #N/A (no bar/label).</summary>
        public double?[] Segments { get; set; }
        public bool IsAnchor { get; set; }
    }

    public sealed class StackedWaterfallResult
    {
        public string[] Categories { get; set; }
        public List<StackedWaterfallRow> Rows { get; set; }
    }

    /// <summary>
    /// Stacked waterfall geometry (FR-C2): the single-waterfall float logic
    /// extended per category. The Base series (shared across categories) floats
    /// each step to the running cumulative of the per-step totals; each category
    /// then stacks its own segment on that floor.
    ///
    /// Anchor (total) rows render as full-height stacked bars (Base = 0). Delta
    /// rows float on the running total.
    ///
    /// Known simplification (to be hardened in W4): mixed-sign categories within
    /// a single step stack in declared order; the UPSLIDE_StackedWaterfall tab
    /// splits positive/negative segments into separate helper series — that
    /// split is deferred. Geometry of Base and anchor detection is final.
    /// </summary>
    public static class StackedWaterfallEngine
    {
        private const double AnchorTolerance = 1e-6;

        public static StackedWaterfallResult Compute(string[] categories, IList<StackedWaterfallPoint> points)
        {
            if (categories == null) throw new ArgumentNullException(nameof(categories));
            if (points == null) throw new ArgumentNullException(nameof(points));

            var rows = new List<StackedWaterfallRow>(points.Count);
            double running = 0;

            for (int i = 0; i < points.Count; i++)
            {
                var p = points[i];
                if (p.Values == null || p.Values.Length != categories.Length)
                    throw new ArgumentException(
                        $"Row '{p.Label}' has {p.Values?.Length ?? 0} values but {categories.Length} categories.");

                double stepTotal = p.Values.Sum();
                bool isAnchor = i == 0 || Math.Abs(stepTotal - running) < AnchorTolerance;

                var segs = new double?[categories.Length];

                if (isAnchor)
                {
                    running = stepTotal;
                    for (int c = 0; c < categories.Length; c++)
                        segs[c] = p.Values[c];

                    rows.Add(new StackedWaterfallRow
                    {
                        Label = p.Label,
                        Base = 0,
                        Segments = segs,
                        IsAnchor = true
                    });
                }
                else
                {
                    // Float the whole step. For a net decrease the floor drops first.
                    double floor = stepTotal >= 0 ? running : running + stepTotal;
                    for (int c = 0; c < categories.Length; c++)
                        segs[c] = p.Values[c] >= 0 ? p.Values[c] : -p.Values[c];

                    running += stepTotal;
                    rows.Add(new StackedWaterfallRow
                    {
                        Label = p.Label,
                        Base = floor,
                        Segments = segs,
                        IsAnchor = false
                    });
                }
            }

            return new StackedWaterfallResult { Categories = categories, Rows = rows };
        }
    }
}
