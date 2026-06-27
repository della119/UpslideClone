using System;
using System.Collections.Generic;
using System.IO;

namespace UpslideClone.Core.Settings
{
    /// <summary>
    /// Per-user settings (FR-S1): theme path + shortcut overrides, persisted under
    /// %APPDATA%\UpslideClone\settings.json. W1 only needs the storage location and
    /// theme path; the editor UI lands in W2/W5. Kept dependency-free.
    /// </summary>
    public sealed class Settings
    {
        public string ThemePath { get; set; }
        public Dictionary<string, string> ShortcutOverrides { get; set; } = new Dictionary<string, string>();

        public static string ConfigDirectory
        {
            get
            {
                var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                return Path.Combine(appData, "UpslideClone");
            }
        }

        public static string ConfigPath => Path.Combine(ConfigDirectory, "settings.json");

        /// <summary>Effective shortcut bindings: defaults overlaid with user overrides.</summary>
        public IDictionary<string, string> EffectiveShortcuts()
        {
            var map = new Dictionary<string, string>(ShortcutMap.Defaults());
            if (ShortcutOverrides != null)
                foreach (var kv in ShortcutOverrides)
                    map[kv.Key] = kv.Value;
            return map;
        }
    }
}
