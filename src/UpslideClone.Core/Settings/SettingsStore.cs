using System;
using System.IO;
using System.Runtime.Serialization.Json;
using System.Text;

namespace UpslideClone.Core.Settings
{
    /// <summary>
    /// Loads/saves per-user <see cref="Settings"/> as JSON under
    /// %APPDATA%\UpslideClone\settings.json (FR-S1). No NuGet — uses
    /// DataContractJsonSerializer (POCO mode). Robust to a missing/corrupt file.
    /// </summary>
    public static class SettingsStore
    {
        public static Settings Load()
        {
            try
            {
                if (!File.Exists(Settings.ConfigPath)) return new Settings();
                using (var fs = File.OpenRead(Settings.ConfigPath))
                    return Parse(fs);
            }
            catch
            {
                return new Settings();
            }
        }

        public static void Save(Settings settings)
        {
            Directory.CreateDirectory(Settings.ConfigDirectory);
            var ser = new DataContractJsonSerializer(typeof(Settings));
            using (var fs = File.Create(Settings.ConfigPath))
                ser.WriteObject(fs, settings ?? new Settings());
        }

        public static Settings Parse(Stream json)
        {
            var ser = new DataContractJsonSerializer(typeof(Settings));
            var s = (Settings)ser.ReadObject(json);
            return s ?? new Settings();
        }

        public static Settings ParseString(string json)
        {
            try
            {
                using (var ms = new MemoryStream(Encoding.UTF8.GetBytes(json ?? "")))
                    return Parse(ms);
            }
            catch
            {
                return new Settings();
            }
        }

        public static string ToJson(Settings settings)
        {
            var ser = new DataContractJsonSerializer(typeof(Settings));
            using (var ms = new MemoryStream())
            {
                ser.WriteObject(ms, settings ?? new Settings());
                return Encoding.UTF8.GetString(ms.ToArray());
            }
        }
    }
}
