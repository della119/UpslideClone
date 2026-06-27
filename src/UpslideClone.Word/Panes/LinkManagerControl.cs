using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Wd = Microsoft.Office.Interop.Word;
using UpslideClone.Core.Util;
using UpslideClone.Word.Linking;

namespace UpslideClone.Word.Panes
{
    /// <summary>Word Link Manager task pane (FR-X3/FR-X4 analogue) — list, refresh, change-source, go-to.</summary>
    public sealed class LinkManagerControl : UserControl
    {
        private readonly ListView _list;
        private readonly Button _btnReload, _btnRefreshAll, _btnChangeSrc, _btnGoTo;

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
            _list.Columns.Add("Type", 56);
            _list.Columns.Add("Source", 130);
            _list.Columns.Add("Sheet!Range", 120);
            _list.Columns.Add("Last refresh", 110);
            _list.Columns.Add("Status", 100);
            _list.DoubleClick += (s, e) => GoToSelected();

            var buttons = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 64, WrapContents = true };
            _btnReload = MakeButton("Reload", (s, e) => Reload());
            _btnRefreshAll = MakeButton("Refresh All", (s, e) => RefreshAll());
            _btnChangeSrc = MakeButton("Change Source…", (s, e) => ChangeSource());
            _btnGoTo = MakeButton("Go To", (s, e) => GoToSelected());
            buttons.Controls.AddRange(new Control[] { _btnReload, _btnRefreshAll, _btnChangeSrc, _btnGoTo });

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

        private Wd.Document ActiveDocument => Globals.ThisAddIn.Application.ActiveDocument;

        public void Reload()
        {
            try
            {
                _list.BeginUpdate();
                _list.Items.Clear();
                foreach (LinkRow row in LinkManagerService.List(ActiveDocument))
                {
                    var item = new ListViewItem(row.Type) { Tag = row.LinkId };
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
                Logger.Error("Word Link Manager reload failed", ex);
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

        private void RefreshAll()
        {
            var result = RefreshEngine.RefreshAll(ActiveDocument);
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
                int n = LinkManagerService.ChangeSource(ActiveDocument, ids, dlg.FileName);
                Reload();
                MessageBox.Show($"Repointed {n} link(s).", "Upslide — Change Source");
            }
        }

        private void GoToSelected()
        {
            if (_list.SelectedItems.Count == 0) return;
            var id = _list.SelectedItems[0].Tag as string;
            if (id != null) LinkManagerService.GoTo(ActiveDocument, id);
        }
    }
}
