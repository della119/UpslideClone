using System.Collections.Generic;
using UpslideClone.Core.Charts;
using Xunit;

namespace UpslideClone.Core.Tests
{
    public class WaterfallEngineTests
    {
        // Fixture: the "Waterfall charts" tab EBITDA bridge 45 -> 32 -> 53.
        // Modelled as: start 45, deltas summing to -13 (down to 32), then deltas
        // up to a 53 total. We assert the float geometry rather than the fixture's
        // exact intermediate split, which the engine derives from running totals.
        [Fact]
        public void FirstRow_IsAlwaysAnchor()
        {
            var pts = new List<WaterfallPoint>
            {
                new WaterfallPoint("EBITDA 2020", 45),
                new WaterfallPoint("Price", 5),
            };

            var rows = WaterfallEngine.Compute(pts);

            Assert.Equal(0, rows[0].Base);
            Assert.Equal(45, rows[0].Total);
            Assert.Null(rows[0].Increase);
            Assert.Null(rows[0].Decrease);
        }

        [Fact]
        public void Increase_FloatsOnRunningTotal()
        {
            var pts = new List<WaterfallPoint>
            {
                new WaterfallPoint("Start", 45),
                new WaterfallPoint("Price", 8),
            };

            var rows = WaterfallEngine.Compute(pts);

            Assert.Equal(45, rows[1].Base);
            Assert.Equal(8, rows[1].Increase);
            Assert.Null(rows[1].Decrease);
            Assert.Null(rows[1].Total);
        }

        [Fact]
        public void Decrease_LowersBaseAndStoresPositiveMagnitude()
        {
            var pts = new List<WaterfallPoint>
            {
                new WaterfallPoint("Start", 45),
                new WaterfallPoint("Cost", -13),
            };

            var rows = WaterfallEngine.Compute(pts);

            // running drops to 32; decrease bar sits on the new floor with height 13.
            Assert.Equal(32, rows[1].Base);
            Assert.Equal(13, rows[1].Decrease);
            Assert.Null(rows[1].Increase);
            Assert.Null(rows[1].Total);
        }

        [Fact]
        public void AnchorDetected_WhenValueEqualsRunningCumulative()
        {
            var pts = new List<WaterfallPoint>
            {
                new WaterfallPoint("EBITDA 2020", 45),
                new WaterfallPoint("Down", -13),   // running -> 32
                new WaterfallPoint("EBITDA mid", 32), // equals running -> anchor
                new WaterfallPoint("Up", 21),       // running -> 53
                new WaterfallPoint("EBITDA 2021", 53), // equals running -> anchor
            };

            var rows = WaterfallEngine.Compute(pts);

            Assert.Equal(32, rows[2].Total);
            Assert.Equal(0, rows[2].Base);
            Assert.Equal(53, rows[4].Total);
            Assert.Equal(0, rows[4].Base);
        }
    }
}
