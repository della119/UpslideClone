namespace UpslideClone.Core.Design
{
    /// <summary>
    /// Pure text helpers for the PowerPoint reference tools (cross-references and
    /// footnotes). The Interop command stores the stable target SlideID on the
    /// shape (<see cref="TargetIdTag"/>) and re-renders the text on refresh.
    /// </summary>
    public static class CrossReference
    {
        /// <summary>Shape tag holding the cross-reference target's stable SlideID.</summary>
        public const string TargetIdTag = "UPS_XRefTargetId";
        /// <summary>Shape tag marking a footnote (value = its number).</summary>
        public const string FootnoteTag = "UPS_Footnote";

        /// <summary>Render a cross-reference line, e.g. "→ Slide 3: Strategy".</summary>
        public static string Format(int slideNumber, string title)
        {
            title = (title ?? "").Trim();
            return title.Length == 0 ? $"→ Slide {slideNumber}" : $"→ Slide {slideNumber}: {title}";
        }

        /// <summary>Render a numbered footnote line, e.g. "1. Source: Annual report".</summary>
        public static string Footnote(int number, string text)
        {
            text = (text ?? "").Trim();
            return $"{number}. {text}";
        }
    }
}
