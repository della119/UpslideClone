using System.Collections.Generic;

namespace UpslideClone.Core.Settings
{
    /// <summary>
    /// Command identifiers shared by the ribbon, the shortcut registrar
    /// (Application.OnKey) and the Settings UI (FR-S2). Names are stable keys.
    /// </summary>
    public static class Commands
    {
        public const string SmartFormat = "SmartFormat";
        public const string BuildWaterfall = "BuildWaterfall";
        public const string BuildStackedWaterfall = "BuildStackedWaterfall";
        public const string DisplayCagr = "DisplayCagr";
        public const string UpdateChart = "UpdateChart";
        public const string ClearFormatting = "ClearFormatting";
        public const string UndoFormatting = "UndoFormatting";

        public const string ToggleTitle = "ToggleTitle";
        public const string ToggleResult = "ToggleResult";
        public const string ToggleItem = "ToggleItem";
        public const string ToggleBackground = "ToggleBackground";
        public const string ToggleBorder = "ToggleBorder";
        public const string ToggleNumberFormat = "ToggleNumberFormat";
        public const string TogglePercentFormat = "TogglePercentFormat";

        // Phase W2+
        public const string Autocolor = "Autocolor";
        public const string ApplyIferror = "ApplyIferror";
        public const string FastFillRight = "FastFillRight";
        public const string FastFillDown = "FastFillDown";
        public const string TrackPrecedent = "TrackPrecedent";
        public const string TrackDependent = "TrackDependent";

        // Phase W3+
        public const string ExportToPowerPoint = "ExportToPowerPoint";
        public const string ExportToWord = "ExportToWord";
        public const string AdvancedExport = "AdvancedExport";
        public const string HighlightLinked = "HighlightLinked";
        public const string ExcelLibrary = "ExcelLibrary";
        public const string SmartPrint = "SmartPrint";
    }

    /// <summary>
    /// Default keyboard-shortcut bindings (Appendix C / training "Keyboard shortcuts" tab),
    /// keyed by command id. Values are Excel <c>Application.OnKey</c> key strings.
    /// Persisted/overridable per-user in W2 (FR-S2).
    /// </summary>
    public static class ShortcutMap
    {
        /// <summary>command id → OnKey key string (e.g. "^+s" = Ctrl+Shift+S).</summary>
        public static IDictionary<string, string> Defaults()
        {
            // ^ = Ctrl, + = Shift, % = Alt (Excel Application.OnKey syntax).
            return new Dictionary<string, string>
            {
                { Commands.SmartFormat,          "^+s" },
                { Commands.BuildWaterfall,       "^+b" },
                { Commands.DisplayCagr,          "^+g" },
                { Commands.UpdateChart,          "^+u" },
                { Commands.ClearFormatting,      "%^c" },

                { Commands.ToggleTitle,          "^+t" },
                { Commands.ToggleResult,         "^+r" },
                { Commands.ToggleItem,           "^+n" },
                { Commands.ToggleBackground,     "^+c" },
                { Commands.ToggleBorder,         "^+o" },
                { Commands.ToggleNumberFormat,   "^+f" },
                { Commands.TogglePercentFormat,  "^+{%}" },

                { Commands.Autocolor,            "^+a" },
                { Commands.ApplyIferror,         "^+i" },
                { Commands.FastFillRight,        "%^r" },
                { Commands.FastFillDown,         "%^d" },
                { Commands.TrackPrecedent,       "^+k" },
                { Commands.TrackDependent,       "^+l" },

                { Commands.ExportToPowerPoint,   "^+e" },
                { Commands.ExportToWord,         "^+d" },
                { Commands.AdvancedExport,       "%^a" },
                { Commands.HighlightLinked,      "%^l" },
                { Commands.ExcelLibrary,         "^+y" },
                { Commands.SmartPrint,           "^+p" },
            };
        }
    }
}
