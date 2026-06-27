using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Json;
using System.Text;

namespace UpslideClone.Core.Library
{
    /// <summary>A reusable saved table/snippet (values only; styling re-applied on insert).</summary>
    public sealed class Snippet
    {
        public string Name { get; set; }
        public int Rows { get; set; }
        public int Cols { get; set; }
        /// <summary>Row-major cell display text.</summary>
        public List<List<string>> Values { get; set; } = new List<List<string>>();
        public long CreatedUtcTicks { get; set; }
    }

    /// <summary>
    /// The Excel Library (FR-L1): a gallery of saved snippets, persisted as JSON
    /// under %APPDATA%\UpslideClone\library.json. Pure model + serialization;
    /// the Excel command reads/writes ranges. No NuGet.
    /// </summary>
    public sealed class SnippetLibrary
    {
        public List<Snippet> Snippets { get; set; } = new List<Snippet>();

        public void Upsert(Snippet s)
        {
            if (s == null) throw new ArgumentNullException(nameof(s));
            Snippets.RemoveAll(x => string.Equals(x.Name, s.Name, StringComparison.OrdinalIgnoreCase));
            Snippets.Add(s);
        }

        public Snippet Get(string name) =>
            Snippets.FirstOrDefault(x => string.Equals(x.Name, name, StringComparison.OrdinalIgnoreCase));

        public bool Remove(string name) =>
            Snippets.RemoveAll(x => string.Equals(x.Name, name, StringComparison.OrdinalIgnoreCase)) > 0;

        public static string DefaultPath
        {
            get
            {
                var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                return Path.Combine(appData, "UpslideClone", "library.json");
            }
        }

        public static SnippetLibrary Load(string path)
        {
            try
            {
                if (string.IsNullOrEmpty(path) || !File.Exists(path)) return new SnippetLibrary();
                using (var fs = File.OpenRead(path))
                    return (SnippetLibrary)new DataContractJsonSerializer(typeof(SnippetLibrary)).ReadObject(fs)
                           ?? new SnippetLibrary();
            }
            catch { return new SnippetLibrary(); }
        }

        public void Save(string path)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            using (var fs = File.Create(path))
                new DataContractJsonSerializer(typeof(SnippetLibrary)).WriteObject(fs, this);
        }

        public string ToJson()
        {
            using (var ms = new MemoryStream())
            {
                new DataContractJsonSerializer(typeof(SnippetLibrary)).WriteObject(ms, this);
                return Encoding.UTF8.GetString(ms.ToArray());
            }
        }

        public static SnippetLibrary FromJson(string json)
        {
            if (string.IsNullOrWhiteSpace(json)) return new SnippetLibrary();
            using (var ms = new MemoryStream(Encoding.UTF8.GetBytes(json)))
                return (SnippetLibrary)new DataContractJsonSerializer(typeof(SnippetLibrary)).ReadObject(ms)
                       ?? new SnippetLibrary();
        }
    }
}
