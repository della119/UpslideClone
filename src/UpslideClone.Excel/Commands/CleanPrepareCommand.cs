using System;
using System.Collections.Generic;
using System.Text;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using UpslideClone.Core.Modelling;

namespace UpslideClone.Excel.Commands
{
    /// <summary>
    /// Clean &amp; Prepare (FR-M6): remove broken defined names (#REF!) before
    /// sharing, and report what was removed. Conservative — leaves valid and
    /// external names in place.
    /// </summary>
    internal static class CleanPrepareCommand
    {
        public static string Run()
        {
            var book = (ExcelInterop.Workbook)ExcelHelpers.App.ActiveWorkbook;
            if (book == null) throw new InvalidOperationException("Open a workbook first.");

            // Collect first, then delete (don't mutate the collection while enumerating).
            var broken = new List<ExcelInterop.Name>();
            foreach (ExcelInterop.Name nm in book.Names)
            {
                string refersTo;
                try { refersTo = nm.RefersTo as string; } catch { continue; }
                if (DefinedNameAudit.IsBroken(refersTo)) broken.Add(nm);
            }

            var report = new StringBuilder();
            int removed = 0;
            foreach (var nm in broken)
            {
                try
                {
                    report.AppendLine("• " + nm.Name);
                    nm.Delete();
                    removed++;
                }
                catch { /* some names refuse deletion; skip */ }
            }

            return removed == 0
                ? "No broken defined names found."
                : $"Removed {removed} broken defined name(s):\n{report}";
        }
    }
}
