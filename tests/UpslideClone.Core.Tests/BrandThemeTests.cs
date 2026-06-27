using UpslideClone.Core.Branding;
using UpslideClone.Core.Util;
using Xunit;

namespace UpslideClone.Core.Tests
{
    public class BrandThemeTests
    {
        [Fact]
        public void Default_HasBrandStandard()
        {
            var t = BrandTheme.Default();
            Assert.Equal("Calibri", t.Fonts.Latin);
            Assert.Equal("华文细黑", t.Fonts.Cjk);
            Assert.Equal("#86BC25", t.Colors.Increase);
            Assert.Contains("EBITDA", t.ResultRowKeywords);
        }

        [Fact]
        public void ParseString_ReadsJson()
        {
            const string json = @"{
                ""fonts"": { ""latin"": ""Arial"", ""cjk"": ""宋体"", ""sizeBody"": 11, ""sizeHeader"": 12 },
                ""colors"": { ""headerFill"": ""#000000"", ""increase"": ""#00FF00"" },
                ""numberFormats"": { ""number"": ""0"", ""percent"": ""0%"", ""currency"": ""$0"" },
                ""resultRowKeywords"": [""WACC""]
            }";

            var t = BrandTheme.ParseString(json);
            Assert.Equal("Arial", t.Fonts.Latin);
            Assert.Equal("#00FF00", t.Colors.Increase);
            Assert.Single(t.ResultRowKeywords);
            Assert.Equal("WACC", t.ResultRowKeywords[0]);
        }

        [Fact]
        public void ColorUtil_FromHex_ParsesRgb()
        {
            var c = ColorUtil.FromHex("#86BC25");
            Assert.Equal(0x86, c.R);
            Assert.Equal(0xBC, c.G);
            Assert.Equal(0x25, c.B);
        }

        [Fact]
        public void ColorUtil_ToOle_IsBgr()
        {
            // OLE/BGR: blue in high byte, red in low byte.
            int ole = ColorUtil.OleFromHex("#FF0000"); // pure red
            Assert.Equal(0x0000FF, ole);
        }
    }
}
