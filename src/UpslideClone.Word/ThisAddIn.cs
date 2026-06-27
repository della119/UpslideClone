using System;
using Office = Microsoft.Office.Core;
using Tools = Microsoft.Office.Tools;
using UpslideClone.Core.Util;
using UpslideClone.Word.Panes;
using UpslideClone.Word.Ribbon;

namespace UpslideClone.Word
{
    /// <summary>
    /// Word VSTO add-in entry point. Hosts the Link Manager task pane and the
    /// Links ribbon group (FR-X3 — refresh side of Excel→Word links).
    /// </summary>
    public partial class ThisAddIn
    {
        private Tools.CustomTaskPane _linkManagerPane;
        private LinkManagerControl _linkManagerControl;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Logger.Info("UpslideClone.Word add-in started.");
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            Logger.Info("UpslideClone.Word add-in shutting down.");
        }

        public void ToggleLinkManager()
        {
            if (_linkManagerPane == null)
            {
                _linkManagerControl = new LinkManagerControl();
                _linkManagerPane = this.CustomTaskPanes.Add(_linkManagerControl, "Upslide Link Manager");
                _linkManagerPane.Width = 360;
            }
            _linkManagerPane.Visible = !_linkManagerPane.Visible;
            if (_linkManagerPane.Visible) _linkManagerControl.Reload();
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new WordRibbon();
        }

        #region VSTO generated code
        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}
