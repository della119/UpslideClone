using System;
using System.IO;
using System.Reflection;
using UpslideClone.Core.Branding;
using UpslideClone.Core.Settings;
using UpslideClone.Core.Util;

namespace UpslideClone.Excel.Commands
{
    /// <summary>
    /// Resolves the active <see cref="BrandTheme"/> (FR-F6). Precedence:
    /// 1) Settings.ThemePath if set, 2) assets\theme.json next to the add-in,
    /// 3) the built-in brand default. Cached; call <see cref="Reload"/> after edits.
    /// </summary>
    internal static class ThemeProvider
    {
        private static BrandTheme _cached;

        public static BrandTheme Current => _cached ?? (_cached = Resolve());

        public static void Reload() => _cached = Resolve();

        private static BrandTheme Resolve()
        {
            try
            {
                var settingsTheme = TryUserThemePath();
                if (settingsTheme != null) return settingsTheme;

                var asmDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? ".";
                foreach (var candidate in new[]
                {
                    Path.Combine(asmDir, "assets", "theme.json"),
                    Path.Combine(asmDir, "theme.json"),
                })
                {
                    if (File.Exists(candidate))
                        return BrandTheme.Load(candidate);
                }
            }
            catch (Exception ex)
            {
                Logger.Warn("Theme resolution failed, using default: " + ex.Message);
            }
            return BrandTheme.Default();
        }

        private static BrandTheme TryUserThemePath()
        {
            try
            {
                if (!File.Exists(Settings.ConfigPath)) return null;
                // W1: settings file is optional; theme path lookup is wired in W2.
                return null;
            }
            catch { return null; }
        }
    }
}
