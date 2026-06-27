using System.Collections.Generic;
using System.Windows.Forms;

namespace UpslideClone.PowerPoint.Design
{
    /// <summary>Tiny WinForms input/picker helpers (no Microsoft.VisualBasic dependency).</summary>
    internal static class Prompt
    {
        public static string Input(string title, string message, string def = "")
        {
            using (var form = new Form { Text = title, FormBorderStyle = FormBorderStyle.FixedDialog, StartPosition = FormStartPosition.CenterScreen, Width = 380, Height = 150, MinimizeBox = false, MaximizeBox = false })
            {
                var lbl = new Label { Text = message, Left = 12, Top = 12, Width = 350 };
                var box = new TextBox { Text = def, Left = 12, Top = 36, Width = 348 };
                var ok = new Button { Text = "OK", DialogResult = DialogResult.OK, Left = 196, Top = 70, Width = 75 };
                var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Left = 285, Top = 70, Width = 75 };
                form.Controls.AddRange(new Control[] { lbl, box, ok, cancel });
                form.AcceptButton = ok; form.CancelButton = cancel;
                return form.ShowDialog() == DialogResult.OK && box.Text.Trim().Length > 0 ? box.Text.Trim() : null;
            }
        }

        public static string Pick(string title, string message, IEnumerable<string> items)
        {
            using (var form = new Form { Text = title, FormBorderStyle = FormBorderStyle.FixedDialog, StartPosition = FormStartPosition.CenterScreen, Width = 380, Height = 320, MinimizeBox = false, MaximizeBox = false })
            {
                var lbl = new Label { Text = message, Left = 12, Top = 12, Width = 350 };
                var list = new ListBox { Left = 12, Top = 36, Width = 348, Height = 200 };
                foreach (var i in items) list.Items.Add(i);
                if (list.Items.Count > 0) list.SelectedIndex = 0;
                var ok = new Button { Text = "OK", DialogResult = DialogResult.OK, Left = 196, Top = 244, Width = 75 };
                var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Left = 285, Top = 244, Width = 75 };
                form.Controls.AddRange(new Control[] { lbl, list, ok, cancel });
                form.AcceptButton = ok; form.CancelButton = cancel;
                return form.ShowDialog() == DialogResult.OK ? list.SelectedItem as string : null;
            }
        }
    }
}
