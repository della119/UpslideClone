using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;

namespace UpslideClone.Core.Branding
{
    [DataContract]
    public sealed class ThemeFonts
    {
        [DataMember(Name = "latin")] public string Latin { get; set; }
        [DataMember(Name = "cjk")] public string Cjk { get; set; }
        [DataMember(Name = "sizeBody")] public double SizeBody { get; set; }
        [DataMember(Name = "sizeHeader")] public double SizeHeader { get; set; }
    }

    [DataContract]
    public sealed class ThemeColors
    {
        [DataMember(Name = "headerFill")] public string HeaderFill { get; set; }
        [DataMember(Name = "headerFont")] public string HeaderFont { get; set; }
        [DataMember(Name = "resultFill")] public string ResultFill { get; set; }
        [DataMember(Name = "border")] public string Border { get; set; }
        [DataMember(Name = "increase")] public string Increase { get; set; }
        [DataMember(Name = "decrease")] public string Decrease { get; set; }
        [DataMember(Name = "total")] public string Total { get; set; }
    }

    [DataContract]
    public sealed class ThemeNumberFormats
    {
        [DataMember(Name = "number")] public string Number { get; set; }
        [DataMember(Name = "percent")] public string Percent { get; set; }
        [DataMember(Name = "currency")] public string Currency { get; set; }
    }

    /// <summary>
    /// Editable branding theme (FR-F6 / NFR-4). Loaded from assets/theme.json so
    /// output can change without recompiling. Defaults match the brand standard.
    /// </summary>
    [DataContract]
    public sealed class BrandTheme
    {
        [DataMember(Name = "fonts")] public ThemeFonts Fonts { get; set; }
        [DataMember(Name = "colors")] public ThemeColors Colors { get; set; }
        [DataMember(Name = "numberFormats")] public ThemeNumberFormats NumberFormats { get; set; }
        [DataMember(Name = "resultRowKeywords")] public List<string> ResultRowKeywords { get; set; }

        /// <summary>The built-in brand default, identical to assets/theme.json shipped with the add-in.</summary>
        public static BrandTheme Default()
        {
            return new BrandTheme
            {
                Fonts = new ThemeFonts { Latin = "Calibri", Cjk = "华文细黑", SizeBody = 10, SizeHeader = 10 },
                Colors = new ThemeColors
                {
                    HeaderFill = "#A9D18E",   // green header (brand/Excel standard)
                    HeaderFont = "#FFFFFF",
                    ResultFill = "#EFEFEF",
                    Border = "#BFBFBF",
                    Increase = "#86BC25",     // bright green deltas up
                    Decrease = "#E0301E",     // red deltas down
                    Total = "#375623"         // dark green anchor/total bars
                },
                NumberFormats = new ThemeNumberFormats
                {
                    Number = "#,##0.0;(#,##0.0)",
                    Percent = "0.0%;(0.0%)",
                    Currency = "¥#,##0.0;(¥#,##0.0)"
                },
                ResultRowKeywords = new List<string>
                {
                    "EBITDA", "EBIT", "Gross Margin", "Net Income", "Total", "毛利", "净利润", "营业利润"
                }
            };
        }

        /// <summary>Load a theme from a JSON file, falling back to <see cref="Default"/> on any error.</summary>
        public static BrandTheme Load(string path)
        {
            try
            {
                if (string.IsNullOrEmpty(path) || !File.Exists(path)) return Default();
                using (var fs = File.OpenRead(path))
                    return Parse(fs);
            }
            catch
            {
                return Default();
            }
        }

        /// <summary>Parse a theme from a JSON stream. Throws on malformed JSON (callers may catch).</summary>
        public static BrandTheme Parse(Stream json)
        {
            var ser = new DataContractJsonSerializer(typeof(BrandTheme));
            var theme = (BrandTheme)ser.ReadObject(json);
            theme.FillDefaults();
            return theme;
        }

        public static BrandTheme ParseString(string json)
        {
            using (var ms = new MemoryStream(Encoding.UTF8.GetBytes(json ?? "")))
                return Parse(ms);
        }

        /// <summary>Backfill any section the JSON omitted so consumers never hit a null sub-object.</summary>
        private void FillDefaults()
        {
            var d = Default();
            if (Fonts == null) Fonts = d.Fonts;
            if (Colors == null) Colors = d.Colors;
            if (NumberFormats == null) NumberFormats = d.NumberFormats;
            if (ResultRowKeywords == null || ResultRowKeywords.Count == 0)
                ResultRowKeywords = d.ResultRowKeywords;
        }
    }
}
