using System.Collections.Generic;
using UpslideClone.Core.Library;
using UpslideClone.Core.Modelling;
using UpslideClone.Core.Settings;
using AppSettings = UpslideClone.Core.Settings.Settings;
using Xunit;

namespace UpslideClone.Core.Tests
{
    public class SettingsStoreTests
    {
        [Fact]
        public void RoundTrips_ThemePathAndOverrides()
        {
            var s = new AppSettings { ThemePath = @"D:\theme.json" };
            s.ShortcutOverrides["SmartFormat"] = "^+q";

            var rebuilt = SettingsStore.ParseString(SettingsStore.ToJson(s));

            Assert.Equal(@"D:\theme.json", rebuilt.ThemePath);
            Assert.Equal("^+q", rebuilt.ShortcutOverrides["SmartFormat"]);
        }

        [Fact]
        public void EffectiveShortcuts_OverlaysDefaults()
        {
            var s = new AppSettings();
            s.ShortcutOverrides["SmartFormat"] = "^+q";
            var eff = s.EffectiveShortcuts();
            Assert.Equal("^+q", eff["SmartFormat"]);          // overridden
            Assert.Equal("^+b", eff[Commands.BuildWaterfall]); // default preserved
        }

        [Fact]
        public void ParseString_Garbage_ReturnsDefaults()
        {
            var s = SettingsStore.ParseString("not json");
            Assert.NotNull(s);
            Assert.Null(s.ThemePath);
        }
    }

    public class SnippetLibraryTests
    {
        private static Snippet Sample(string name) => new Snippet
        {
            Name = name,
            Rows = 2,
            Cols = 2,
            Values = new List<List<string>>
            {
                new List<string> { "Revenue", "100" },
                new List<string> { "EBITDA", "32" }
            }
        };

        [Fact]
        public void Upsert_ReplacesByName_CaseInsensitive()
        {
            var lib = new SnippetLibrary();
            lib.Upsert(Sample("IS"));
            lib.Upsert(Sample("is")); // same name (case-insensitive) → replace
            Assert.Single(lib.Snippets);
        }

        [Fact]
        public void ToJson_FromJson_RoundTrips()
        {
            var lib = new SnippetLibrary();
            lib.Upsert(Sample("IS"));
            var rebuilt = SnippetLibrary.FromJson(lib.ToJson());

            var s = rebuilt.Get("IS");
            Assert.NotNull(s);
            Assert.Equal(2, s.Rows);
            Assert.Equal("EBITDA", s.Values[1][0]);
        }

        [Fact]
        public void Remove_DeletesSnippet()
        {
            var lib = new SnippetLibrary();
            lib.Upsert(Sample("IS"));
            Assert.True(lib.Remove("IS"));
            Assert.Empty(lib.Snippets);
        }
    }

    public class DefinedNameAuditTests
    {
        [Theory]
        [InlineData("=Sheet1!#REF!", true)]
        [InlineData("=#REF!$A$1", true)]
        [InlineData("=Sheet1!$A$1", false)]
        [InlineData("", false)]
        public void IsBroken_DetectsRefErrors(string refersTo, bool expected)
        {
            Assert.Equal(expected, DefinedNameAudit.IsBroken(refersTo));
        }

        [Theory]
        [InlineData("=[Book2.xlsx]Sheet1!$A$1", true)]
        [InlineData(@"='C:\models\[deal.xlsx]IS'!$A$1", true)]
        [InlineData("=Sheet1!$A$1", false)]
        public void IsExternal_DetectsLinks(string refersTo, bool expected)
        {
            Assert.Equal(expected, DefinedNameAudit.IsExternal(refersTo));
        }
    }
}
