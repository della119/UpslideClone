using System.Collections.Generic;
using UpslideClone.Core.Design;
using Xunit;

namespace UpslideClone.Core.Tests
{
    public class AlignEngineTests
    {
        private static List<LayoutBox> Boxes() => new List<LayoutBox>
        {
            new LayoutBox(10, 10, 100, 20),
            new LayoutBox(40, 50, 60, 40),
            new LayoutBox(200, 5, 80, 30),
        };

        [Fact]
        public void Align_Left_SetsAllToMinLeft()
        {
            var r = AlignEngine.Align(Boxes(), AlignMode.Left);
            Assert.All(r, b => Assert.Equal(10, b.Left));
        }

        [Fact]
        public void Align_Right_SetsRightEdgesEqual()
        {
            var r = AlignEngine.Align(Boxes(), AlignMode.Right);
            float maxR = 280; // 200+80
            Assert.All(r, b => Assert.Equal(maxR, b.Right, 3f));
        }

        [Fact]
        public void Align_CenterHorizontal_CentersOnGroupMidline()
        {
            var r = AlignEngine.Align(Boxes(), AlignMode.CenterHorizontal);
            float cx = (10 + 280) / 2f;
            Assert.All(r, b => Assert.Equal(cx, b.CenterX, 3f));
        }

        [Fact]
        public void Align_PreservesSize()
        {
            var r = AlignEngine.Align(Boxes(), AlignMode.Top);
            Assert.Equal(100, r[0].Width);
            Assert.Equal(40, r[1].Height);
        }

        [Fact]
        public void Distribute_Horizontal_EqualGaps()
        {
            var boxes = new List<LayoutBox>
            {
                new LayoutBox(0, 0, 10, 10),
                new LayoutBox(15, 0, 10, 10),
                new LayoutBox(100, 0, 10, 10),
            };
            var r = AlignEngine.Distribute(boxes, DistributeAxis.Horizontal);
            float gap1 = r[1].Left - r[0].Right;
            float gap2 = r[2].Left - r[1].Right;
            Assert.Equal(gap1, gap2, 3f);
            Assert.Equal(0, r[0].Left, 3f);     // ends fixed
            Assert.Equal(100, r[2].Left, 3f);
        }

        [Fact]
        public void Distribute_TwoBoxes_Unchanged()
        {
            var boxes = new List<LayoutBox> { new LayoutBox(0, 0, 10, 10), new LayoutBox(100, 0, 10, 10) };
            var r = AlignEngine.Distribute(boxes, DistributeAxis.Horizontal);
            Assert.Equal(0, r[0].Left);
            Assert.Equal(100, r[1].Left);
        }

        [Fact]
        public void MatchSize_Both_ResizesToFirst()
        {
            var r = AlignEngine.MatchSize(Boxes(), SizeMatch.Both);
            Assert.Equal(100, r[1].Width);
            Assert.Equal(20, r[1].Height);
            Assert.Equal(100, r[0].Width); // reference unchanged
        }
    }

    public class SlideCheckRulesTests
    {
        [Fact]
        public void IsOffSlide_DetectsOverflow()
        {
            Assert.True(SlideCheckRules.IsOffSlide(new LayoutBox(-5, 10, 100, 20), 960, 540));
            Assert.True(SlideCheckRules.IsOffSlide(new LayoutBox(900, 10, 100, 20), 960, 540));
            Assert.False(SlideCheckRules.IsOffSlide(new LayoutBox(10, 10, 100, 20), 960, 540));
        }

        [Fact]
        public void IsTinyText_BelowFloor()
        {
            Assert.True(SlideCheckRules.IsTinyText(8));
            Assert.False(SlideCheckRules.IsTinyText(12));
        }
    }

    public class CrossReferenceTests
    {
        [Fact]
        public void Format_WithTitle()
        {
            Assert.Equal("→ Slide 3: Strategy", CrossReference.Format(3, "Strategy"));
        }

        [Fact]
        public void Format_NoTitle_OmitsColon()
        {
            Assert.Equal("→ Slide 5", CrossReference.Format(5, "  "));
        }

        [Fact]
        public void Footnote_Numbered()
        {
            Assert.Equal("1. Source: Annual report", CrossReference.Footnote(1, "Source: Annual report"));
        }
    }

    public class TableOfContentsTests
    {
        [Fact]
        public void Build_SkipsUntitledAndSkipped()
        {
            var titles = new List<KeyValuePair<int, string>>
            {
                new KeyValuePair<int, string>(1, "Cover"),
                new KeyValuePair<int, string>(2, "  "),       // untitled
                new KeyValuePair<int, string>(3, "Strategy"),
                new KeyValuePair<int, string>(4, "Financials"),
            };
            var toc = TableOfContents.Build(titles, new HashSet<int> { 1 });
            Assert.Equal(2, toc.Count);
            Assert.Equal("Strategy", toc[0].Title);
            Assert.Equal(3, toc[0].SlideIndex);
        }

        [Fact]
        public void Render_NumbersEntries()
        {
            var toc = new List<TocEntry> { new TocEntry { Title = "A" }, new TocEntry { Title = "B" } };
            var text = TableOfContents.Render(toc);
            Assert.Contains("1.\tA", text);
            Assert.Contains("2.\tB", text);
        }
    }
}
