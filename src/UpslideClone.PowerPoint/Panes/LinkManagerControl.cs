using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using PPT = Microsoft.Office.Interop.PowerPoint;
using UpslideClone.Core.Util;
using UpslideClone.PowerPoint.Linking;

namespace UpslideClone.PowerPoint.Panes
{
    /// <summary>
    /// Link Manager task-pane UI (FR-X4): lists every linked object with its
    /// source/status, and drives Refresh / Change Source / Go To. Built in code
    /// (no designer/resx) to keep the project shell minimal.
    /// </summary>
    public sealed class LinkManagerControl : UserControl
    {
        private readonly ListView _list;
        private readonly Button _btnReload, _btnRefreshSel, _btnRefreshAll, _btnChangeSrc, _btnGoTo, _btnCheck;

        public LinkManagerControl()
        {
            _list = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                MultiSelect = true,
                HideSelection = false,
                GridLines = true
            };
            _list.Columns.Add("Slide", 44);
            _list.Columns.Add("Type", 56);
            _list.Columns.Add("Source", 120);
            _list.Columns.Add("Sheet!Range", 110);
            _list.Columns.Add("Last refresh", 110);
            _list.Columns.Add("Status", 90);
            _list.DoubleClick += (s, e) => GoToSelected();

            var buttons = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 64, WrapContents = true };
            _btnReload = MakeButton("Reload", (s, e) => Reload());
            _btnCheck = MakeButton("Check Sources", (s, e) => Reload(true));
            _btnRefreshSel = MakeButton("Refresh Selected", (s, e) => RefreshSelected());
            _btnRefreshAll = MakeButton("Refresh All", (s, e) => RefreshAll());
            _btnChangeSrc = MakeButton("Change Source…", (s, e) => ChangeSource());
            _btnGoTo = MakeButton("Go To", (s, e) => GoToSelected());
            buttons.Controls.AddRange(new Control[] { _btnReload, _btnCheck, _btnRefreshSel, _btnRefreshAll, _btnChangeSrc, _btnGoTo });

            Controls.Add(_list);
            Controls.Add(buttons);
            MinimumSize = new Size(300, 200);
        }

        private static Button MakeButton(string text, EventHandler onClick)
        {
            var b = new Button { Text = text, AutoSize = true, Margin = new Padding(3) };
            b.Click += onClick;
            return b;
        }

        private PPT.Presentation ActivePresentation => Globals.ThisAddIn.Application.ActivePresentation;
        private PPT.Application App => Globals.ThisAddIn.Application;

        public void Reload() => Reload(false);

        public void Reload(bool checkDrift)
        {
            try
            {
                _list.BeginUpdate();
                _list.Items.Clear();
                foreach (LinkRow row in LinkManagerService.List(ActivePresentation, checkDrift))
                {
                    var item = new ListViewItem(row.SlideIndex.ToString()) { Tag = row.LinkId };
                    item.SubItems.Add(row.Type);
                    item.SubItems.Add(row.SourceFile);
                    item.SubItems.Add(row.SheetRange);
                    item.SubItems.Add(row.LastRefresh);
                    item.SubItems.Add(row.Status);
                    item.ToolTipText = row.SourcePath;
                    if (row.Status != "OK") item.ForeColor = Color.Firebrick;
                    _list.Items.Add(item);
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Link Manager reload failed", ex);
                MessageBox.Show(ex.Message, "Upslide — Link Manager");
            }
            finally { _list.EndUpdate(); }
        }

        private HashSet<string> SelectedIds()
        {
            var ids = new HashSet<string>(StringComparer.Ordinal);
            foreach (ListViewItem it in _list.SelectedItems)
                if (it.Tag is string id) ids.Add(id);
            return ids;
        }

        private void RefreshSelected()
        {
            var ids = SelectedIds();
            if (ids.Count == 0) { MessageBox.Show("Select one or more links first."); return; }
            var result = RefreshEngine.RefreshLinks(ActivePresentation, ids);
            Reload();
            MessageBox.Show(result.ToString(), "Upslide — Refresh");
        }

        private void RefreshAll()
        {
            var result = RefreshEngine.RefreshAll(ActivePresentation);
            Reload();
            MessageBox.Show(result.ToString(), "Upslide — Refresh All");
        }

        private void ChangeSource()
        {
            var ids = SelectedIds();
            if (ids.Count == 0) { MessageBox.Show("Select the links to repoint first."); return; }
            using (var dlg = new OpenFileDialog { Filter = "Excel workbooks|*.xlsx;*.xlsm;*.xls|All files|*.*" })
            {
                if (dlg.ShowDialog() != DialogResult.OK) return;
                int n = LinkManagerService.ChangeSource(ActivePresentation, ids, dlg.FileName);
                Reload();
                MessageBox.Show($"Repointed {n} link(s) to {dlg.FileName}.", "Upslide — Change Source");
            }
        }

        private void GoToSelected()
        {
            if (_list.SelectedItems.Count == 0) return;
            var id = _list.SelectedItems[0].Tag as string;
            if (id != null) LinkManagerService.GoTo(App, ActivePresentation, id);
        }
    }
}
