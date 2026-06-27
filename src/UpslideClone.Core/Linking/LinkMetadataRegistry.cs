using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;

namespace UpslideClone.Core.Linking
{
    /// <summary>
    /// A document-level registry of full <see cref="LinkMetadata"/> records, keyed
    /// by LinkId. Used by the Word add-in (FR-X3): Word shapes/tables have no
    /// Shape.Tags collection like PowerPoint, so link metadata is stored in a
    /// document CustomXMLPart and each linked object is anchored by a bookmark
    /// named <see cref="BookmarkPrefix"/> + LinkId. Pure + unit-testable.
    /// </summary>
    public sealed class LinkMetadataRegistry
    {
        public const string NamespaceUri = "urn:upslideclone:doclinks";
        public const string BookmarkPrefix = "UPS_";

        public List<LinkMetadata> Items { get; set; } = new List<LinkMetadata>();

        public static string BookmarkName(string linkId) => BookmarkPrefix + linkId;

        public void Upsert(LinkMetadata m)
        {
            if (m == null) throw new ArgumentNullException(nameof(m));
            Items.RemoveAll(i => i.LinkId == m.LinkId);
            Items.Add(m);
        }

        public LinkMetadata Get(string linkId) => Items.FirstOrDefault(i => i.LinkId == linkId);

        public string ToXml()
        {
            XNamespace ns = NamespaceUri;
            var root = new XElement(ns + "links",
                Items.Select(m => new XElement(ns + "link",
                    new XAttribute("id", m.LinkId ?? ""),
                    new XAttribute("workbook", m.SourceWorkbook ?? ""),
                    new XAttribute("sheet", m.SourceSheet ?? ""),
                    new XAttribute("range", m.SourceRange ?? ""),
                    new XAttribute("type", m.ExportType.ToString()),
                    new XAttribute("hash", m.SourceHash ?? ""),
                    new XAttribute("lastRefresh", m.LastRefresh?.ToString("o", CultureInfo.InvariantCulture) ?? ""))));
            return root.ToString(SaveOptions.DisableFormatting);
        }

        public static LinkMetadataRegistry FromXml(string xml)
        {
            var reg = new LinkMetadataRegistry();
            if (string.IsNullOrWhiteSpace(xml)) return reg;

            XNamespace ns = NamespaceUri;
            XElement root;
            try { root = XElement.Parse(xml); }
            catch { return reg; }

            foreach (var e in root.Elements(ns + "link"))
            {
                ExportType et;
                Enum.TryParse((string)e.Attribute("type"), out et);
                DateTime lr;
                DateTime? last = DateTime.TryParse((string)e.Attribute("lastRefresh"),
                    CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out lr) ? lr : (DateTime?)null;

                reg.Items.Add(new LinkMetadata
                {
                    LinkId = (string)e.Attribute("id"),
                    SourceWorkbook = (string)e.Attribute("workbook"),
                    SourceSheet = (string)e.Attribute("sheet"),
                    SourceRange = (string)e.Attribute("range"),
                    ExportType = et,
                    SourceHash = (string)e.Attribute("hash"),
                    LastRefresh = last
                });
            }
            return reg;
        }
    }
}
