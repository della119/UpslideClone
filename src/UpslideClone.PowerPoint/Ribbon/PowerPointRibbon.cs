using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using UpslideClone.Core.Util;
using UpslideClone.Core.Design;
using UpslideClone.PowerPoint.Design;
using UpslideClone.PowerPoint.Linking;

namespace UpslideClone.PowerPoint.Ribbon
{
    [System.Runtime.InteropServices.ComVisible(true)]
    public sealed class PowerPointRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbon;

        public string GetCustomUI(string ribbonID)
        {
            return ReadResource("UpslideClone.PowerPoint.Ribbon.PowerPointRibbon.xml");
        }

        public void OnRibbonLoad(Office.IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
            Logger.Info("Upslide PowerPoint ribbon loaded.");
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

        public void OnInsertPlaceholder(Office.IRibbonControl c)
        {
            try
            {
                var msg = SizingGuideCommand.InsertPlaceholder(Globals.ThisAddIn.Application);
                MessageBox.Show(msg, "Upslide — Sizing Guide", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                Logger.Error("InsertPlaceholder failed", ex);
                MessageBox.Show(ex.Message, "Upslide — Sizing Guide", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        public void OnRefreshAll(Office.IRibbonControl c)
        {
            try
            {
                var result = RefreshEngine.RefreshAll(Globals.ThisAddIn.Application.ActivePresentation);
                MessageBox.Show(result.ToString(), "Upslide — Refresh All", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                Logger.Error("RefreshAll failed", ex);
                MessageBox.Show(ex.Message, "Upslide — Refresh All", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        // ---- Design toolkit ----
        public void OnAlignLeft(Office.IRibbonControl c) => Run("Align Left", () => DesignCommands.Align(AlignMode.Left));
        public void OnAlignCenterH(Office.IRibbonControl c) => Run("Align Center", () => DesignCommands.Align(AlignMode.CenterHorizontal));
        public void OnAlignRight(Office.IRibbonControl c) => Run("Align Right", () => DesignCommands.Align(AlignMode.Right));
        public void OnAlignTop(Office.IRibbonControl c) => Run("Align Top", () => DesignCommands.Align(AlignMode.Top));
        public void OnAlignMiddle(Office.IRibbonControl c) => Run("Align Middle", () => DesignCommands.Align(AlignMode.Middle));
        public void OnAlignBottom(Office.IRibbonControl c) => Run("Align Bottom", () => DesignCommands.Align(AlignMode.Bottom));
        public void OnDistributeH(Office.IRibbonControl c) => Run("Distribute Across", () => DesignCommands.Distribute(DistributeAxis.Horizontal));
        public void OnDistributeV(Office.IRibbonControl c) => Run("Distribute Down", () => DesignCommands.Distribute(DistributeAxis.Vertical));
        public void OnSameSize(Office.IRibbonControl c) => Run("Same Size", () => DesignCommands.MatchSize(SizeMatch.Both));
        public void OnBringFront(Office.IRibbonControl c) => Run("Bring to Front", () => DesignCommands.ZOrder(true));
        public void OnSendBack(Office.IRibbonControl c) => Run("Send to Back", () => DesignCommands.ZOrder(false));
        public void OnGroup(Office.IRibbonControl c) => Run("Group", () => DesignCommands.Group(true));
        public void OnUngroup(Office.IRibbonControl c) => Run("Ungroup", () => DesignCommands.Group(false));
        public void OnFormatShapes(Office.IRibbonControl c) => Run("Format Shapes", () => DesignCommands.FormatShapes());
        public void OnSelectSimilar(Office.IRibbonControl c) => Run("Select Similar", () => DesignCommands.SelectSimilar());
        public void OnSmartPainter(Office.IRibbonControl c) => Run("Smart Painter", () => DesignCommands.SmartPainter());
        public void OnToc(Office.IRibbonControl c) => Run("Table of Contents", () => DesignCommands.BuildTableOfContents());
        public void OnSlideCheck(Office.IRibbonControl c) => Run("Slide Check", () => DesignCommands.SlideCheck());
        public void OnOutline(Office.IRibbonControl c) { try { Globals.ThisAddIn.ToggleOutline(); } catch (Exception ex) { Logger.Error("Outline", ex); MessageBox.Show(ex.Message, "Upslide — Outline"); } }

        // ---- Content Library (#3) + References (#2) ----
        public void OnSaveContent(Office.IRibbonControl c) => Run("Save to Library", () => ContentLibrary.Save(Globals.ThisAddIn.Application));
        public void OnInsertContent(Office.IRibbonControl c) => Run("Insert from Library", () => ContentLibrary.Insert(Globals.ThisAddIn.Application));
        public void OnFootnote(Office.IRibbonControl c) => Run("Insert Footnote", () => ReferenceCommands.InsertFootnote(Globals.ThisAddIn.Application));
        public void OnXref(Office.IRibbonControl c) => Run("Insert Cross-reference", () => ReferenceCommands.InsertCrossReference(Globals.ThisAddIn.Application));
        public void OnXrefRefresh(Office.IRibbonControl c) => Run("Refresh Cross-references", () => ReferenceCommands.RefreshCrossReferences(Globals.ThisAddIn.Application));

        /// <summary>Wrap a design command: run, log, surface a friendly message on failure.</summary>
        private void Run(string name, Func<string> action)
        {
            try
            {
                var status = action();
                Logger.Info($"PPT command ok: {name} — {status}");
            }
            catch (Exception ex)
            {
                Logger.Error($"PPT command failed: {name}", ex);
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
