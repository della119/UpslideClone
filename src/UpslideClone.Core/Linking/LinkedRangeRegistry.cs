using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace UpslideClone.Core.Linking
{
    /// <summary>A source range that has been exported/linked out (for FR-X10 highlight).</summary>
    public sealed class LinkedRange
    {
        public string Sheet { get; set; }
        public string Range { get; set; }
        public string LinkId { get; set; }
    }

    /// <summary>
    /// The registry of linked-out source ranges, mirrored in the source workbook
    /// as a CustomXMLPart (§6.4) so "Highlight linked items" (FR-X10) can flag
    /// which ranges feed a deck. Pure serialization — unit-testable.
    /// </summary>
    public sealed class LinkedRangeRegistry
    {
        public const string NamespaceUri = "urn:upslideclone:links";

        public List<LinkedRange> Items { get; set; } = new List<LinkedRange>();

        /// <summary>Add or update by LinkId (idempotent re-export).</summary>
        public void Upsert(string linkId, string sheet, string range)
        {
            var existing = Items.FirstOrDefault(i => i.LinkId == linkId);
            if (existing != null) { existing.Sheet = sheet; existing.Range = range; }
            else Items.Add(new LinkedRange { LinkId = linkId, Sheet = sheet, Range = range });
        }

        public string ToXml()
        {
            XNamespace ns = NamespaceUri;
            var root = new XElement(ns + "links",
                Items.Select(i => new XElement(ns + "link",
                    new XAttribute("id", i.LinkId ?? ""),
                    new XAttribute("sheet", i.Sheet ?? ""),
                    new XAttribute("range", i.Range ?? ""))));
            return root.ToString(SaveOptions.DisableFormatting);
        }

        public static LinkedRangeRegistry FromXml(string xml)
        {
            var reg = new LinkedRangeRegistry();
            if (string.IsNullOrWhiteSpace(xml)) return reg;

            XNamespace ns = NamespaceUri;
            XElement root;
            try { root = XElement.Parse(xml); }
            catch { return reg; }

            foreach (var e in root.Elements(ns + "link"))
            {
                reg.Items.Add(new LinkedRange
                {
                    LinkId = (string)e.Attribute("id"),
                    Sheet = (string)e.Attribute("sheet"),
                    Range = (string)e.Attribute("range")
                });
            }
            return reg;
        }
    }
}
