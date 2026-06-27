using System;
using System.Collections.Generic;
using System.Linq;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using UpslideClone.Core.Library;

namespace UpslideClone.Excel.Commands
{
    /// <summary>Excel Library (FR-L1): save the selection as a reusable snippet; insert one back.</summary>
    internal static class ExcelLibraryCommand
    {
        public static string SaveSelection()
        {
            var sel = ExcelHelpers.RequireRangeSelection();
            string name = Prompt.Input("Save to Library", "Snippet name:");
            if (name == null) return "Cancelled.";

            int rows = sel.Rows.Count, cols = sel.Columns.Count;
            var snippet = new Snippet { Name = name, Rows = rows, Cols = cols, CreatedUtcTicks = DateTime.UtcNow.Ticks };
            for (int r = 1; r <= rows; r++)
            {
                var rowList = new List<string>(cols);
                for (int c = 1; c <= cols; c++)
                {
                    var cell = (ExcelInterop.Range)sel.Cells[r, c];
                    rowList.Add(cell.Text == null ? "" : cell.Text.ToString());
                }
                snippet.Values.Add(rowList);
            }

            var lib = SnippetLibrary.Load(SnippetLibrary.DefaultPath);
            lib.Upsert(snippet);
            lib.Save(SnippetLibrary.DefaultPath);
            return $"Saved \"{name}\" ({rows}×{cols}) to the library.";
        }

        public static string InsertFromLibrary()
        {
            var lib = SnippetLibrary.Load(SnippetLibrary.DefaultPath);
            if (lib.Snippets.Count == 0) return "The library is empty — save a selection first.";

            string name = Prompt.Pick("Insert from Library", "Choose a snippet:", lib.Snippets.Select(s => s.Name));
            if (name == null) return "Cancelled.";
            var snippet = lib.Get(name);
            if (snippet == null) return "Snippet not found.";

            var sel = ExcelHelpers.RequireRangeSelection();
            var sheet = (ExcelInterop.Worksheet)sel.Worksheet;
            int top = sel.Row, left = sel.Column;

            var arr = new object[snippet.Rows, snippet.Cols];
            for (int r = 0; r < snippet.Rows; r++)
                for (int c = 0; c < snippet.Cols; c++)
                    arr[r, c] = r < snippet.Values.Count && c < snippet.Values[r].Count ? snippet.Values[r][c] : "";

            var dest = ExcelHelpers.Sub(sheet, top, left, snippet.Rows, snippet.Cols);
            dest.Value2 = arr;
            return $"Inserted \"{name}\" ({snippet.Rows}×{snippet.Cols}).";
        }
    }
}
