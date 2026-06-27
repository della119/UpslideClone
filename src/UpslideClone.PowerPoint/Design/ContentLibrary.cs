using System;
using System.IO;
using System.Linq;
using Office = Microsoft.Office.Core;
using PPT = Microsoft.Office.Interop.PowerPoint;

namespace UpslideClone.PowerPoint.Design
{
    /// <summary>
    /// Content Library (#3): save selected shapes as reusable items and insert them
    /// later. Items are stored as slides in a background library presentation
    /// (%APPDATA%\UpslideClone\ppt-library.pptx), one slide per item, named by the
    /// item name — PowerPoint's own copy/paste preserves full fidelity.
    /// </summary>
    internal static class ContentLibrary
    {
        private const int ppLayoutBlank = 12;

        private static string LibPath
        {
            get
            {
                // PowerPoint refuses to SaveAs into %APPDATA%\Roaming, so the library
                // lives under Documents (a valid Office save location).
                var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "UpslideClone");
                Directory.CreateDirectory(dir);
                return Path.Combine(dir, "ppt-library.pptx");
            }
        }

        // Opened WITH a window: PowerPoint's clipboard Paste only populates a
        // presentation that has a document window.
        private static PPT.Presentation OpenLib(PPT.Application app, bool createIfMissing)
        {
            if (File.Exists(LibPath))
                return app.Presentations.Open(LibPath);
            if (!createIfMissing) return null;
            var p = app.Presentations.Add();
            p.SaveAs(LibPath);
            return p;
        }

        public static string Save(PPT.Application app)
        {
            var userWindow = app.ActiveWindow;
            var sel = userWindow.Selection;
            if (sel.Type != PPT.PpSelectionType.ppSelectionShapes)
                throw new InvalidOperationException("Select one or more shapes to save first.");

            string name = Prompt.Input("Save to Library", "Item name:");
            if (name == null) return "Cancelled.";

            sel.ShapeRange.Copy();
            PPT.Presentation lib = OpenLib(app, true);
            try
            {
                var slide = lib.Slides.Add(lib.Slides.Count + 1, (PPT.PpSlideLayout)ppLayoutBlank);
                slide.Shapes.Paste();
                try { slide.Name = UniqueName(lib, name); } catch { /* name constraints — ignore */ }
                lib.Save();
            }
            finally { lib.Close(); try { userWindow.Activate(); } catch { } }
            return $"Saved \"{name}\" to the library.";
        }

        public static string Insert(PPT.Application app)
        {
            if (!File.Exists(LibPath)) return "The library is empty — save a selection first.";
            var userWindow = app.ActiveWindow;
            var target = (PPT.Slide)userWindow.View.Slide;

            PPT.Presentation lib = OpenLib(app, false);
            string pick;
            try
            {
                var names = Enumerable.Range(1, lib.Slides.Count).Select(i => lib.Slides[i].Name).ToList();
                if (names.Count == 0) return "The library is empty.";
                pick = Prompt.Pick("Insert from Library", "Choose an item:", names);
                if (pick == null) return "Cancelled.";

                PPT.Slide src = null;
                for (int i = 1; i <= lib.Slides.Count; i++)
                    if (lib.Slides[i].Name == pick) { src = lib.Slides[i]; break; }
                if (src == null) return "Item not found.";
                src.Shapes.Range(Type.Missing).Copy();
            }
            finally { lib.Close(); }

            try { userWindow.Activate(); } catch { }
            target.Shapes.Paste();
            return $"Inserted \"{pick}\".";
        }

        private static string UniqueName(PPT.Presentation lib, string name)
        {
            var existing = Enumerable.Range(1, lib.Slides.Count - 1).Select(i => lib.Slides[i].Name).ToList();
            string candidate = name;
            int n = 2;
            while (existing.Contains(candidate)) candidate = $"{name} ({n++})";
            return candidate;
        }
    }
}
