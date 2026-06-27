using System;
using UpslideClone.Core.Linking;
using Xunit;

namespace UpslideClone.Core.Tests
{
    public class LinkMetadataRegistryTests
    {
        private static LinkMetadata Sample(string id, ExportType type) => new LinkMetadata
        {
            LinkId = id,
            SourceWorkbook = @"D:\m\deal.xlsx",
            SourceSheet = "IS",
            SourceRange = "$B$2:$F$20",
            ExportType = type,
            SourceHash = "h" + id,
            LastRefresh = new DateTime(2026, 6, 17, 10, 0, 0, DateTimeKind.Utc)
        };

        [Fact]
        public void Upsert_ReplacesByLinkId()
        {
            var reg = new LinkMetadataRegistry();
            reg.Upsert(Sample("a", ExportType.Picture));
            reg.Upsert(Sample("a", ExportType.Table)); // same id → replace
            reg.Upsert(Sample("b", ExportType.Picture));
            Assert.Equal(2, reg.Items.Count);
            Assert.Equal(ExportType.Table, reg.Get("a").ExportType);
        }

        [Fact]
        public void ToXml_FromXml_RoundTrips()
        {
            var reg = new LinkMetadataRegistry();
            reg.Upsert(Sample("a", ExportType.Table));
            var rebuilt = LinkMetadataRegistry.FromXml(reg.ToXml());

            var m = rebuilt.Get("a");
            Assert.NotNull(m);
            Assert.Equal(@"D:\m\deal.xlsx", m.SourceWorkbook);
            Assert.Equal("IS", m.SourceSheet);
            Assert.Equal("$B$2:$F$20", m.SourceRange);
            Assert.Equal(ExportType.Table, m.ExportType);
            Assert.Equal("ha", m.SourceHash);
            Assert.Equal(Sample("a", ExportType.Table).LastRefresh, m.LastRefresh);
        }

        [Fact]
        public void BookmarkName_PrefixesLinkId()
        {
            Assert.Equal("UPS_abc123", LinkMetadataRegistry.BookmarkName("abc123"));
        }

        [Fact]
        public void FromXml_Garbage_ReturnsEmpty()
        {
            Assert.Empty(LinkMetadataRegistry.FromXml("<nope").Items);
        }
    }
}
