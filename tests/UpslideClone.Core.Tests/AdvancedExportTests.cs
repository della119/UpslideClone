using System;
using UpslideClone.Core.Linking;
using Xunit;

namespace UpslideClone.Core.Tests
{
    public class AdvancedExportMapTests
    {
        [Fact]
        public void Parse_ReadsEntries_ByHeaderKeyword()
        {
            // 1-based grid like Excel Value2.
            var grid = new object[4, 4];
            grid[0, 0] = "Source Sheet"; grid[0, 1] = "Range"; grid[0, 2] = "Slide"; grid[0, 3] = "Type";
            grid[1, 0] = "WACC"; grid[1, 1] = "B2:F20"; grid[1, 2] = 2; grid[1, 3] = "Table";
            grid[2, 0] = "WACC"; grid[2, 1] = "B22:F30"; grid[2, 2] = 3; grid[2, 3] = "Picture";
            grid[3, 0] = ""; grid[3, 1] = ""; grid[3, 2] = ""; grid[3, 3] = ""; // blank → skipped

            var entries = AdvancedExportMap.Parse(grid);

            Assert.Equal(2, entries.Count);
            Assert.Equal("WACC", entries[0].SourceSheet);
            Assert.Equal("B2:F20", entries[0].SourceRange);
            Assert.Equal(2, entries[0].TargetSlide);
            Assert.Equal(ExportType.Table, entries[0].ExportType);
            Assert.Equal(ExportType.Picture, entries[1].ExportType);
        }

        [Fact]
        public void Parse_DefaultsTypeToPicture_AndSlideToZero()
        {
            var grid = new object[2, 1];
            grid[0, 0] = "Range";
            grid[1, 0] = "A1:B2";
            var entries = AdvancedExportMap.Parse(grid);
            Assert.Single(entries);
            Assert.Equal(ExportType.Picture, entries[0].ExportType);
            Assert.Equal(0, entries[0].TargetSlide);
        }

        [Fact]
        public void Parse_Throws_WhenNoRangeColumn()
        {
            var grid = new object[2, 2];
            grid[0, 0] = "Sheet"; grid[0, 1] = "Slide";
            Assert.Throws<FormatException>(() => AdvancedExportMap.Parse(grid));
        }
    }

    public class LinkedRangeRegistryTests
    {
        [Fact]
        public void Upsert_AddsThenUpdatesByLinkId()
        {
            var reg = new LinkedRangeRegistry();
            reg.Upsert("id1", "Sheet1", "A1:B2");
            reg.Upsert("id1", "Sheet1", "A1:C5"); // update same id
            reg.Upsert("id2", "Sheet2", "D1");
            Assert.Equal(2, reg.Items.Count);
            Assert.Equal("A1:C5", reg.Items[0].Range);
        }

        [Fact]
        public void ToXml_FromXml_RoundTrips()
        {
            var reg = new LinkedRangeRegistry();
            reg.Upsert("id1", "Income Statement", "B2:F20");
            reg.Upsert("id2", "WACC", "B2:F30");

            var rebuilt = LinkedRangeRegistry.FromXml(reg.ToXml());

            Assert.Equal(2, rebuilt.Items.Count);
            Assert.Equal("id1", rebuilt.Items[0].LinkId);
            Assert.Equal("Income Statement", rebuilt.Items[0].Sheet);
            Assert.Equal("B2:F20", rebuilt.Items[0].Range);
        }

        [Fact]
        public void FromXml_Empty_ReturnsEmptyRegistry()
        {
            Assert.Empty(LinkedRangeRegistry.FromXml("").Items);
            Assert.Empty(LinkedRangeRegistry.FromXml(null).Items);
            Assert.Empty(LinkedRangeRegistry.FromXml("<garbage").Items);
        }
    }
}
