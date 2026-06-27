using System;
using System.Collections.Generic;
using System.Globalization;

namespace UpslideClone.Core.Linking
{
    /// <summary>How an exported object was materialised in the target document.</summary>
    public enum ExportType
    {
        Picture,
        Table,
        Chart,
        Text
    }

    /// <summary>
    /// The link tag stored on each exported PowerPoint/Word shape (via Shape.Tags).
    /// This is the W3 data model — defined now so Core stays the single source of
    /// truth for the link schema (the UPS_* tag keys live in <see cref="TagKeys"/>).
    /// </summary>
    public sealed class LinkMetadata
    {
        public string LinkId { get; set; }
        public string SourceWorkbook { get; set; }
        public string SourceSheet { get; set; }
        public string SourceRange { get; set; }
        public ExportType ExportType { get; set; }
        public DateTime? LastRefresh { get; set; }
        public string SourceHash { get; set; }

        public static string NewId() => Guid.NewGuid().ToString("N");

        private const string IsoRoundTrip = "o";

        /// <summary>Serialise to the UPS_* key/value pairs stored on a PowerPoint/Word Shape.Tags collection.</summary>
        public IDictionary<string, string> ToTags()
        {
            var t = new Dictionary<string, string>
            {
                { TagKeys.LinkId, LinkId ?? "" },
                { TagKeys.SourceWorkbook, SourceWorkbook ?? "" },
                { TagKeys.SourceSheet, SourceSheet ?? "" },
                { TagKeys.SourceRange, SourceRange ?? "" },
                { TagKeys.ExportType, ExportType.ToString() },
                { TagKeys.SourceHash, SourceHash ?? "" },
                { TagKeys.LastRefresh, LastRefresh?.ToString(IsoRoundTrip, CultureInfo.InvariantCulture) ?? "" },
            };
            return t;
        }

        /// <summary>Rebuild from a tag lookup (e.g. a function reading Shape.Tags by key). Returns null if untagged.</summary>
        public static LinkMetadata FromTags(Func<string, string> get)
        {
            if (get == null) throw new ArgumentNullException(nameof(get));
            string id = get(TagKeys.LinkId);
            if (string.IsNullOrEmpty(id)) return null; // not one of our shapes

            var m = new LinkMetadata
            {
                LinkId = id,
                SourceWorkbook = NullIfEmpty(get(TagKeys.SourceWorkbook)),
                SourceSheet = NullIfEmpty(get(TagKeys.SourceSheet)),
                SourceRange = NullIfEmpty(get(TagKeys.SourceRange)),
                SourceHash = NullIfEmpty(get(TagKeys.SourceHash)),
            };

            ExportType et;
            m.ExportType = Enum.TryParse(get(TagKeys.ExportType), out et) ? et : ExportType.Picture;

            string lr = get(TagKeys.LastRefresh);
            DateTime parsed;
            if (!string.IsNullOrEmpty(lr) &&
                DateTime.TryParse(lr, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out parsed))
                m.LastRefresh = parsed;

            return m;
        }

        private static string NullIfEmpty(string s) => string.IsNullOrEmpty(s) ? null : s;
    }

    /// <summary>Native Shape.Tags keys used to persist <see cref="LinkMetadata"/> on a shape.</summary>
    public static class TagKeys
    {
        public const string LinkId = "UPS_LinkId";
        public const string SourceWorkbook = "UPS_SourceWorkbook";
        public const string SourceSheet = "UPS_SourceSheet";
        public const string SourceRange = "UPS_SourceRange";
        public const string ExportType = "UPS_ExportType";
        public const string LastRefresh = "UPS_LastRefresh";
        public const string SourceHash = "UPS_SourceHash";
        public const string Placeholder = "UPS_Placeholder";
    }
}
