using System;
using System.IO;
using System.Text;

namespace UpslideClone.Core.Util
{
    /// <summary>
    /// Minimal rolling file logger (NFR-8). Writes to
    /// %APPDATA%\UpslideClone\logs\upslide-yyyyMMdd.log. UI-free so it lives in Core;
    /// the add-in projects route every ribbon callback's try/catch through here.
    /// </summary>
    public static class Logger
    {
        public enum Level { Debug, Info, Warn, Error }

        private static readonly object Gate = new object();

        /// <summary>Minimum level actually written. Configurable via Settings later.</summary>
        public static Level MinLevel = Level.Info;

        public static string LogDirectory
        {
            get
            {
                var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                return Path.Combine(appData, "UpslideClone", "logs");
            }
        }

        public static void Info(string message) => Write(Level.Info, message, null);
        public static void Warn(string message) => Write(Level.Warn, message, null);
        public static void Debug(string message) => Write(Level.Debug, message, null);
        public static void Error(string message, Exception ex = null) => Write(Level.Error, message, ex);

        public static void Write(Level level, string message, Exception ex)
        {
            if (level < MinLevel) return;
            try
            {
                Directory.CreateDirectory(LogDirectory);
                var file = Path.Combine(LogDirectory, "upslide-" + DateTime.Now.ToString("yyyyMMdd") + ".log");

                var sb = new StringBuilder();
                sb.Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"))
                  .Append(" [").Append(level).Append("] ")
                  .Append(message);
                if (ex != null) sb.AppendLine().Append(ex);

                lock (Gate)
                    File.AppendAllText(file, sb.ToString() + Environment.NewLine, Encoding.UTF8);
            }
            catch
            {
                // Logging must never throw into a ribbon callback.
            }
        }
    }
}
