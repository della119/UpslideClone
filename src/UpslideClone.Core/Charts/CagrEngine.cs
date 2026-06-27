using System;
using System.Collections.Generic;
using System.Linq;

namespace UpslideClone.Core.Charts
{
    /// <summary>Result of a CAGR computation plus the data needed to draw the arrow annotation.</summary>
    public sealed class CagrResult
    {
        public double FirstValue { get; set; }
        public double LastValue { get; set; }
        /// <summary>Number of compounding periods (intervals) between first and last point.</summary>
        public int Periods { get; set; }
        /// <summary>Compound annual growth rate as a fraction (e.g. 0.312 = 31.2%).</summary>
        public double Cagr { get; set; }
        /// <summary>Pre-formatted label, e.g. "CAGR +31.2%".</summary>
        public string Label { get; set; }
    }

    /// <summary>
    /// CAGR arrow computation (FR-C4) — pure, unit-testable.
    /// CAGR = (last / first)^(1 / periods) - 1, where periods = number of
    /// intervals between the first and last data point.
    /// </summary>
    public static class CagrEngine
    {
        public static double Cagr(double first, double last, int periods)
        {
            if (periods <= 0) throw new ArgumentOutOfRangeException(nameof(periods), "periods must be >= 1.");
            if (first <= 0) throw new ArgumentOutOfRangeException(nameof(first), "first value must be > 0 for CAGR.");
            return Math.Pow(last / first, 1.0 / periods) - 1.0;
        }

        /// <summary>Compute CAGR over a series of points (uses first and last; periods = count - 1).</summary>
        public static CagrResult Compute(IList<double> series)
        {
            if (series == null) throw new ArgumentNullException(nameof(series));
            var nums = series.Where(v => !double.IsNaN(v)).ToList();
            if (nums.Count < 2)
                throw new ArgumentException("Need at least 2 points to compute CAGR.");

            double first = nums[0];
            double last = nums[nums.Count - 1];
            int periods = nums.Count - 1;
            double cagr = Cagr(first, last, periods);

            return new CagrResult
            {
                FirstValue = first,
                LastValue = last,
                Periods = periods,
                Cagr = cagr,
                Label = "CAGR " + (cagr >= 0 ? "+" : "") + (cagr).ToString("0.0%", System.Globalization.CultureInfo.InvariantCulture)
            };
        }
    }
}
