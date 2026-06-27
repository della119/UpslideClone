using System;
using System.Collections.Generic;
using UpslideClone.Core.Linking;
using Xunit;

namespace UpslideClone.Core.Tests
{
    public class LinkHashTests
    {
        [Fact]
        public void Compute_IsDeterministic()
        {
            var a = new object[,] { { "Rev", 100.0 }, { "EBITDA", 32.0 } };
            var b = new object[,] { { "Rev", 100.0 }, { "EBITDA", 32.0 } };
            Assert.Equal(LinkHash.Compute(a), LinkHash.Compute(b));
        }

        [Fact]
        public void Compute_ChangesWhenDataChanges()
        {
            var a = new object[,] { { "Rev", 100.0 } };
            var b = new object[,] { { "Rev", 101.0 } };
            Assert.NotEqual(LinkHash.Compute(a), LinkHash.Compute(b));
        }

        [Fact]
        public void Compute_NullDiffersFromEmptyString()
        {
            var a = new object[,] { { "x", null } };
            var b = new object[,] { { "x", "" } };
            Assert.NotEqual(LinkHash.Compute(a), LinkHash.Compute(b));
        }

        [Fact]
        public void Compute_ProducesSha256HexLength()
        {
            string h = LinkHash.Compute("hello");
            Assert.Equal(64, h.Length); // 32 bytes → 64 hex chars
        }
    }

    public class LinkMetadataTests
    {
        [Fact]
        public void ToTags_FromTags_RoundTrips()
        {
            var original = new LinkMetadata
            {
                LinkId = LinkMetadata.NewId(),
                SourceWorkbook = @"D:\models\deal.xlsx",
                SourceSheet = "Income Statement",
                SourceRange = "B2:F20",
                ExportType = ExportType.Table,
                SourceHash = "abc123",
                LastRefresh = new DateTime(2026, 6, 17, 9, 30, 0, DateTimeKind.Utc)
            };

            var tags = original.ToTags();
            var rebuilt = LinkMetadata.FromTags(k => tags.TryGetValue(k, out var v) ? v : null);

            Assert.NotNull(rebuilt);
            Assert.Equal(original.LinkId, rebuilt.LinkId);
            Assert.Equal(original.SourceWorkbook, rebuilt.SourceWorkbook);
            Assert.Equal(original.SourceSheet, rebuilt.SourceSheet);
            Assert.Equal(original.SourceRange, rebuilt.SourceRange);
            Assert.Equal(ExportType.Table, rebuilt.ExportType);
            Assert.Equal(original.SourceHash, rebuilt.SourceHash);
            Assert.Equal(original.LastRefresh, rebuilt.LastRefresh);
        }

        [Fact]
        public void FromTags_ReturnsNull_WhenNoLinkId()
        {
            var empty = new Dictionary<string, string>();
            Assert.Null(LinkMetadata.FromTags(k => empty.TryGetValue(k, out var v) ? v : null));
        }

        [Fact]
        public void FromTags_DefaultsExportTypeToPicture_OnGarbage()
        {
            var tags = new Dictionary<string, string>
            {
                { TagKeys.LinkId, "id1" },
                { TagKeys.ExportType, "NotAType" }
            };
            var m = LinkMetadata.FromTags(k => tags.TryGetValue(k, out var v) ? v : null);
            Assert.Equal(ExportType.Picture, m.ExportType);
        }
    }
}
