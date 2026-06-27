using System.Collections.Generic;
using System.Linq;

namespace UpslideClone.Core.Design
{
    public sealed class TocEntry
    {
        public int SlideIndex { get; set; }
        public string Title { get; set; }
    }

    /// <summary>
    /// Pure Table-of-Contents builder: turn (slide index, title) pairs into TOC
    /// entries, skipping untitled slides and an optional cover/divider set.
    /// </summary>
    public static class TableOfContents
    {
        /// <param name="slideTitles">(1-based slide index, raw title) for every slide.</param>
        /// <param name="skipIndices">Slides to omit (e.g. the cover or the TOC slide itself).</param>
        public static List<TocEntry> Build(IEnumerable<KeyValuePair<int, string>> slideTitles, ISet<int> skipIndices = null)
        {
            var entries = new List<TocEntry>();
            foreach (var kv in slideTitles)
            {
                if (skipIndices != null && skipIndices.Contains(kv.Key)) continue;
                string title = (kv.Value ?? "").Trim();
                if (title.Length == 0) continue;        // untitled → not in the TOC
                entries.Add(new TocEntry { SlideIndex = kv.Key, Title = title });
            }
            return entries;
        }

        /// <summary>Render the TOC as plain lines ("1.  Title") for a text box.</summary>
        public static string Render(IEnumerable<TocEntry> entries)
        {
            return string.Join("\r", entries.Select((e, i) => $"{i + 1}.\t{e.Title}"));
        }
    }
}
