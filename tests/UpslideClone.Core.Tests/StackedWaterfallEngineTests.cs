using System.Collections.Generic;
using UpslideClone.Core.Charts;
using Xunit;

namespace UpslideClone.Core.Tests
{
    public class StackedWaterfallEngineTests
    {
        // "Stacked Waterfall" tab: EBITDA bridge split FR/UK/DE.
        private static readonly string[] Cats = { "FR", "UK", "DE" };

        [Fact]
        public void AnchorRow_HasZeroBaseAndFullSegments()
        {
            var pts = new List<StackedWaterfallPoint>
            {
                new StackedWaterfallPoint("EBITDA 2020", new double[] { 20, 15, 10 }), // total 45, anchor
            };

            var res = StackedWaterfallEngine.Compute(Cats, pts);
            var row = res.Rows[0];

            Assert.True(row.IsAnchor);
            Assert.Equal(0, row.Base);
            Assert.Equal(20, row.Segments[0]);
            Assert.Equal(15, row.Segments[1]);
            Assert.Equal(10, row.Segments[2]);
        }

        [Fact]
        public void PositiveStep_FloatsOnRunningTotal()
        {
            var pts = new List<StackedWaterfallPoint>
            {
                new StackedWaterfallPoint("EBITDA 2020", new double[] { 20, 15, 10 }), // 45 anchor
                new StackedWaterfallPoint("Price",       new double[] { 3, 2, 1 }),    // +6 delta
            };

            var res = StackedWaterfallEngine.Compute(Cats, pts);
            var delta = res.Rows[1];

            Assert.False(delta.IsAnchor);
            Assert.Equal(45, delta.Base);
            Assert.Equal(3, delta.Segments[0]);
            Assert.Equal(2, delta.Segments[1]);
            Assert.Equal(1, delta.Segments[2]);
        }

        [Fact]
        public void NegativeStep_DropsFloor()
        {
            var pts = new List<StackedWaterfallPoint>
            {
                new StackedWaterfallPoint("EBITDA 2020", new double[] { 20, 15, 10 }), // 45 anchor
                new StackedWaterfallPoint("Cost",        new double[] { -5, -4, -4 }), // -13 delta
            };

            var res = StackedWaterfallEngine.Compute(Cats, pts);
            var delta = res.Rows[1];

            Assert.Equal(32, delta.Base);          // floor drops by 13
            Assert.Equal(5, delta.Segments[0]);    // magnitude stored positive
            Assert.Equal(4, delta.Segments[1]);
            Assert.Equal(4, delta.Segments[2]);
        }
    }
}
