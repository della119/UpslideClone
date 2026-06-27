using System;
using System.Collections.Generic;
using UpslideClone.Core.Charts;
using Xunit;

namespace UpslideClone.Core.Tests
{
    public class CagrEngineTests
    {
        [Fact]
        public void Cagr_NetSales_9_to_15_5_over_2_periods()
        {
            // "CAGR Arrow" fixture: net sales 9 -> 15.5, 2020..2022 = 2 intervals.
            double cagr = CagrEngine.Cagr(9.0, 15.5, 2);
            Assert.Equal(0.3123, cagr, 4);
        }

        [Fact]
        public void Compute_FromSeries_UsesFirstAndLast()
        {
            var r = CagrEngine.Compute(new List<double> { 9.0, 12.0, 15.5 });
            Assert.Equal(9.0, r.FirstValue);
            Assert.Equal(15.5, r.LastValue);
            Assert.Equal(2, r.Periods);
            Assert.StartsWith("CAGR +31", r.Label);
        }

        [Fact]
        public void Compute_NegativeGrowth_LabelHasNoPlus()
        {
            var r = CagrEngine.Compute(new List<double> { 100.0, 81.0 });
            Assert.True(r.Cagr < 0);
            Assert.StartsWith("CAGR -", r.Label);
        }

        [Fact]
        public void Cagr_Throws_OnNonPositiveFirst()
        {
            Assert.Throws<ArgumentOutOfRangeException>(() => CagrEngine.Cagr(0, 10, 2));
        }
    }
}
