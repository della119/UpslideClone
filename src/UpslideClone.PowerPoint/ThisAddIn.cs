using System;
using Office = Microsoft.Office.Core;
using Tools = Microsoft.Office.Tools;
using UpslideClone.Core.Util;
using UpslideClone.PowerPoint.Panes;
using UpslideClone.PowerPoint.Ribbon;

namespace UpslideClone.PowerPoint
{
    /// <summary>
    /// PowerPoint VSTO add-in entry point. Hosts the Link Manager task pane
    /// (FR-X4) and serves the Links ribbon group.
    /// </summary>
    public partial class ThisAddIn
    {
        private Tools.CustomTaskPane _linkManagerPane;
        private LinkManagerControl _linkManagerControl;
        private Tools.CustomTaskPane _outlinePane;
        private OutlineControl _outlineControl;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Logger.Info("UpslideClone.PowerPoint add-in started.");
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            Logger.Info("UpslideClone.PowerPoint add-in shutting down.");
        }

        /// <summary>Create (once) and toggle the Link Manager task pane.</summary>
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

        /// <summary>Create (once) and toggle the Outline task pane (#2).</summary>
        public void ToggleOutline()
        {
            if (_outlinePane == null)
            {
                _outlineControl = new OutlineControl();
                _outlinePane = this.CustomTaskPanes.Add(_outlineControl, "Upslide Outline");
                _outlinePane.Width = 320;
            }
            _outlinePane.Visible = !_outlinePane.Visible;
            if (_outlinePane.Visible) _outlineControl.Reload();
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new PowerPointRibbon();
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
