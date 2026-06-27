using System.Collections.Generic;

namespace UpslideClone.Core.Design
{
    /// <summary>A single issue found by Slide Check (FR Slide Check / Finalize).</summary>
    public sealed class SlideIssue
    {
        public int SlideIndex { get; set; }
        public string Kind { get; set; }     // e.g. "OffSlide", "NoTitle", "TinyText"
        public string Detail { get; set; }
        public override string ToString() => $"Slide {SlideIndex}: {Kind} — {Detail}";
    }

    /// <summary>
    /// Pure Slide Check rules (no Office dependency). The PowerPoint command feeds
    /// shape geometry / titles in and surfaces the returned issues.
    /// </summary>
    public static class SlideCheckRules
    {
        /// <summary>A shape that extends beyond the slide canvas (with a small tolerance).</summary>
        public static bool IsOffSlide(LayoutBox shape, float slideWidth, float slideHeight, float tolerance = 1f)
        {
            return shape.Left < -tolerance
                || shape.Top < -tolerance
                || shape.Right > slideWidth + tolerance
                || shape.Bottom > slideHeight + tolerance;
        }

        /// <summary>True if a text size is below the legibility floor (default 10pt).</summary>
        public static bool IsTinyText(double fontSize, double floor = 10.0)
        {
            return fontSize > 0 && fontSize < floor;
        }

        /// <summary>Build the off-slide issue list for one slide's shapes.</summary>
        public static IEnumerable<SlideIssue> CheckOffSlide(int slideIndex, IEnumerable<KeyValuePair<string, LayoutBox>> namedShapes, float slideWidth, float slideHeight)
        {
            foreach (var kv in namedShapes)
                if (IsOffSlide(kv.Value, slideWidth, slideHeight))
                    yield return new SlideIssue { SlideIndex = slideIndex, Kind = "OffSlide", Detail = kv.Key };
        }
    }
}
