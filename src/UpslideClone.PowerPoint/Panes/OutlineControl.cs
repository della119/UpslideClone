using System;
using System.Drawing;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PPT = Microsoft.Office.Interop.PowerPoint;
using UpslideClone.Core.Util;

namespace UpslideClone.PowerPoint.Panes
{
    /// <summary>Outline pane (#2): lists slide titles; double-click jumps to a slide.</summary>
    public sealed class OutlineControl : UserControl
    {
        private readonly ListView _list;

        public OutlineControl()
        {
            _list = new ListView { Dock = DockStyle.Fill, View = View.Details, FullRowSelect = true, HideSelection = false, GridLines = true };
            _list.Columns.Add("#", 36);
            _list.Columns.Add("Title", 280);
            _list.DoubleClick += (s, e) => GoTo();

            var bar = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 36 };
            var reload = new Button { Text = "Reload", AutoSize = true, Margin = new Padding(3) };
            reload.Click += (s, e) => Reload();
            bar.Controls.Add(reload);

            Controls.Add(_list);
            Controls.Add(bar);
            MinimumSize = new Size(260, 200);
        }

        private PPT.Application App => Globals.ThisAddIn.Application;

        public void Reload()
        {
            try
            {
                _list.BeginUpdate();
                _list.Items.Clear();
                var pres = App.ActivePresentation;
                for (int i = 1; i <= pres.Slides.Count; i++)
                {
                    string title = "";
                    try { if (pres.Slides[i].Shapes.HasTitle == Office.MsoTriState.msoTrue) title = pres.Slides[i].Shapes.Title.TextFrame.TextRange.Text; }
                    catch { }
                    var item = new ListViewItem(i.ToString()) { Tag = i };
                    item.SubItems.Add(string.IsNullOrWhiteSpace(title) ? "(untitled)" : title.Replace("\r", " ").Replace("\v", " "));
                    if (string.IsNullOrWhiteSpace(title)) item.ForeColor = Color.Gray;
                    _list.Items.Add(item);
                }
            }
            catch (Exception ex) { Logger.Error("Outline reload failed", ex); }
            finally { _list.EndUpdate(); }
        }

        private void GoTo()
        {
            if (_list.SelectedItems.Count == 0) return;
            if (_list.SelectedItems[0].Tag is int idx)
            {
                try { App.ActiveWindow.View.GotoSlide(idx); } catch { }
            }
        }
    }
}
