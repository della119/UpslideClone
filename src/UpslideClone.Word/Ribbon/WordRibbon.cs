using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using UpslideClone.Core.Util;
using UpslideClone.Word.Formatting;
using UpslideClone.Word.Linking;

namespace UpslideClone.Word.Ribbon
{
    [System.Runtime.InteropServices.ComVisible(true)]
    public sealed class WordRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbon;

        public string GetCustomUI(string ribbonID)
        {
            return ReadResource("UpslideClone.Word.Ribbon.WordRibbon.xml");
        }

        public void OnRibbonLoad(Office.IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
            Logger.Info("Upslide Word ribbon loaded.");
        }

        public void OnToggleLinkManager(Office.IRibbonControl c)
        {
            try { Globals.ThisAddIn.ToggleLinkManager(); }
            catch (Exception ex)
            {
                Logger.Error("ToggleLinkManager failed", ex);
                MessageBox.Show(ex.Message, "Upslide — Link Manager", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        public void OnRefreshAll(Office.IRibbonControl c)
        {
            try
            {
                var doc = Globals.ThisAddIn.Application.ActiveDocument;
                var result = RefreshEngine.RefreshAll(doc);
                MessageBox.Show(result.ToString(), "Upslide — Refresh All", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                Logger.Error("RefreshAll failed", ex);
                MessageBox.Show(ex.Message, "Upslide — Refresh All", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        public void OnFormatTable(Office.IRibbonControl c) => Run("Format Table", () => FormatCommands.FormatTable(Globals.ThisAddIn.Application));
        public void OnBrandFont(Office.IRibbonControl c) => Run("Brand Font", () => FormatCommands.ApplyBrandFont(Globals.ThisAddIn.Application));

        private void Run(string name, Func<string> action)
        {
            try
            {
                var status = action();
                Logger.Info($"Word command ok: {name} — {status}");
            }
            catch (Exception ex)
            {
                Logger.Error($"Word command failed: {name}", ex);
                MessageBox.Show(ex.Message, "Upslide — " + name, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private static string ReadResource(string resourceName)
        {
            var asm = Assembly.GetExecutingAssembly();
            using (var stream = asm.GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                    throw new InvalidOperationException("Embedded ribbon XML not found: " + resourceName);
                using (var reader = new StreamReader(stream))
                    return reader.ReadToEnd();
            }
        }
    }
}
