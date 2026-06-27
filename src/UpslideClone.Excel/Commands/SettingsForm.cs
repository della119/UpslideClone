using System;
using System.Drawing;
using System.Windows.Forms;
using UpslideClone.Core.Settings;
using CoreSettings = UpslideClone.Core.Settings.Settings;

namespace UpslideClone.Excel.Commands
{
    /// <summary>
    /// Settings dialog (FR-S1/FR-S2): edit the branding theme path and the keyboard
    /// shortcut map; persisted per-user via <see cref="SettingsStore"/>.
    /// </summary>
    internal sealed class SettingsForm : Form
    {
        private readonly TextBox _themePath;
        private readonly DataGridView _grid;

        public SettingsForm()
        {
            Text = "Upslide Settings";
            StartPosition = FormStartPosition.CenterScreen;
            Width = 540;
            Height = 480;
            MinimizeBox = false;
            MaximizeBox = false;

            var current = SettingsStore.Load();

            var themeLbl = new Label { Text = "Branding theme (theme.json) — blank = built-in brand default:", Left = 12, Top = 12, Width = 500 };
            _themePath = new TextBox { Left = 12, Top = 34, Width = 410, Text = current.ThemePath ?? "" };
            var browse = new Button { Text = "Browse…", Left = 428, Top = 32, Width = 84 };
            browse.Click += (s, e) =>
            {
                using (var dlg = new OpenFileDialog { Filter = "Theme JSON|*.json|All files|*.*" })
                    if (dlg.ShowDialog() == DialogResult.OK) _themePath.Text = dlg.FileName;
            };

            var shortLbl = new Label { Text = "Keyboard shortcuts (Excel OnKey syntax: ^=Ctrl, +=Shift, %=Alt):", Left = 12, Top = 66, Width = 500 };
            _grid = new DataGridView
            {
                Left = 12,
                Top = 88,
                Width = 500,
                Height = 300,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                RowHeadersVisible = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            };
            _grid.Columns.Add("cmd", "Command");
            _grid.Columns.Add("key", "Shortcut");
            _grid.Columns["cmd"].ReadOnly = true;
            foreach (var kv in current.EffectiveShortcuts())
                _grid.Rows.Add(kv.Key, kv.Value);

            var save = new Button { Text = "Save", DialogResult = DialogResult.OK, Left = 348, Top = 398, Width = 75 };
            var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Left = 437, Top = 398, Width = 75 };
            save.Click += (s, e) => Persist();

            Controls.AddRange(new Control[] { themeLbl, _themePath, browse, shortLbl, _grid, save, cancel });
            AcceptButton = save;
            CancelButton = cancel;
        }

        private void Persist()
        {
            var s = SettingsStore.Load();
            s.ThemePath = string.IsNullOrWhiteSpace(_themePath.Text) ? null : _themePath.Text.Trim();

            s.ShortcutOverrides.Clear();
            var defaults = ShortcutMap.Defaults();
            foreach (DataGridViewRow row in _grid.Rows)
            {
                string cmd = row.Cells["cmd"].Value as string;
                string key = row.Cells["key"].Value as string;
                if (string.IsNullOrEmpty(cmd)) continue;
                string def;
                if (!defaults.TryGetValue(cmd, out def) || !string.Equals(def, key, StringComparison.Ordinal))
                    s.ShortcutOverrides[cmd] = key ?? "";
            }

            SettingsStore.Save(s);
            ThemeProvider.Reload();
        }
    }

    internal static class SettingsCommand
    {
        public static string Run()
        {
            using (var form = new SettingsForm())
                return form.ShowDialog() == DialogResult.OK ? "Settings saved." : "Settings unchanged.";
        }
    }
}
