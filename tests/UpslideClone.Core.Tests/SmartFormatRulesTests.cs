using System.Collections.Generic;
using UpslideClone.Core.Formatting;
using Xunit;

namespace UpslideClone.Core.Tests
{
    public class SmartFormatRulesTests
    {
        [Theory]
        [InlineData("EBITDA", true)]
        [InlineData("Gross Margin", true)]
        [InlineData("Net Income", true)]
        [InlineData("Total", true)]
        [InlineData("毛利", true)]
        [InlineData("净利润", true)]
        [InlineData("营业利润", true)]
        [InlineData("合计", true)]
        [InlineData("Revenue", false)]
        [InlineData("COGS", false)]
        [InlineData("", false)]
        public void IsResultRow_MatchesKeywords(string label, bool expected)
        {
            Assert.Equal(expected, SmartFormatRules.IsResultRow(label));
        }

        [Fact]
        public void IsPercentColumn_TrueOnHeaderKeyword()
        {
            Assert.True(SmartFormatRules.IsPercentColumn("CAGR", new List<double> { 5, 10 }));
            Assert.True(SmartFormatRules.IsPercentColumn("Margin %", new List<double> { 100 }));
        }

        [Fact]
        public void IsPercentColumn_TrueWhenAllRatios()
        {
            Assert.True(SmartFormatRules.IsPercentColumn("col", new List<double> { 0.31, 0.18, 0.5 }));
        }

        [Fact]
        public void IsPercentColumn_FalseWhenIntegersOrLargeValues()
        {
            Assert.False(SmartFormatRules.IsPercentColumn("col", new List<double> { 1, 2, 3 }));
            Assert.False(SmartFormatRules.IsPercentColumn("col", new List<double> { 45, 32, 53 }));
        }

        [Fact]
        public void IsPercentColumn_FalseWhenEmpty()
        {
            Assert.False(SmartFormatRules.IsPercentColumn("col", new List<double>()));
        }

        [Fact]
        public void IsResultRow_UsesThemeKeywordsWhenProvided()
        {
            var keywords = new[] { "WACC", "营业利润" };
            Assert.True(SmartFormatRules.IsResultRow("WACC optimum", keywords));
            Assert.False(SmartFormatRules.IsResultRow("EBITDA", keywords)); // not in custom set
        }
    }
}
