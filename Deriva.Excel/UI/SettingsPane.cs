using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Deriva.Excel.ETTJ;

namespace Deriva.Excel.UI
{
    internal sealed class SettingsPane : UserControl
    {
        private readonly TextBox _holidayUrl;
        private readonly CheckedListBox _curves;
        private readonly TextBox _cacheDir;

        internal SettingsPane()
        {
            Dock = DockStyle.Fill;
            BackColor = Color.White;
            AutoScaleMode = AutoScaleMode.Dpi;

            var settings = EttjSettings.Load();
            var selectedCurves = LoadSelectedCurves(settings);

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(24),
                ColumnCount = 1,
                RowCount = 10,
                AutoScroll = true
            };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Controls.Add(root);

            root.Controls.Add(CreateTitle("Settings"));

            root.Controls.Add(CreateLabel("Holiday fetch URL"));
            _holidayUrl = new TextBox
            {
                Text = settings.HolidayUrl,
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 0, 0, 14)
            };
            root.Controls.Add(_holidayUrl);

            root.Controls.Add(CreateLabel("ETTJ curves fetched on Excel init"));
            _curves = new CheckedListBox
            {
                Dock = DockStyle.Fill,
                CheckOnClick = true,
                IntegralHeight = false,
                MinimumSize = new Size(0, 260),
                Margin = new Padding(0, 0, 0, 14)
            };
            foreach (var pair in EttjCurveCatalog.AvailableCurves.OrderBy(p => p.Key, StringComparer.Ordinal))
            {
                var item = new CurveItem(pair.Key, pair.Value);
                _curves.Items.Add(item, selectedCurves.Contains(pair.Key));
            }
            root.Controls.Add(_curves);

            root.Controls.Add(CreateLabel("ETTJ cache directory"));
            _cacheDir = new TextBox
            {
                Text = settings.EttjCacheDir,
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 0, 0, 14)
            };
            root.Controls.Add(_cacheDir);

            root.Controls.Add(CreateReadonlyLine("Source", EttjSettings.B3Source));
            root.Controls.Add(CreateReadonlyLine("B3 base URL", EttjSettings.B3BaseUrlTemplate));

            var save = new Button
            {
                Text = "Save",
                AutoSize = true,
                MinimumSize = new Size(120, 38),
                Margin = new Padding(0, 18, 0, 0),
                Padding = new Padding(16, 6, 16, 6)
            };
            save.Click += (sender, args) => SaveSettings();
            root.Controls.Add(save);
        }

        private void SaveSettings()
        {
            var selected = _curves.CheckedItems
                .Cast<CurveItem>()
                .Select(item => item.Code)
                .ToList();

            if (selected.Count == 0)
            {
                MessageBox.Show(
                    "Select at least one ETTJ curve.",
                    "Deriva Settings",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            var settings = new EttjSettings
            {
                HolidayUrl = _holidayUrl.Text,
                InitCurves = string.Join(",", selected.ToArray()),
                EttjCacheDir = _cacheDir.Text
            };

            try
            {
                settings.Save();
                EttjMemoryCache.Clear();
                MessageBox.Show(
                    "Settings saved.",
                    "Deriva Settings",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Unable to save settings: " + ex.Message,
                    "Deriva Settings",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private static HashSet<string> LoadSelectedCurves(EttjSettings settings)
        {
            try
            {
                return new HashSet<string>(
                    EttjService.ParseCurveCsv(settings.InitCurves),
                    StringComparer.OrdinalIgnoreCase);
            }
            catch
            {
                return new HashSet<string>(
                    new[] { "PRE", "DIC" },
                    StringComparer.OrdinalIgnoreCase);
            }
        }

        private static Label CreateTitle(string text)
        {
            return new Label
            {
                Text = text,
                Dock = DockStyle.Fill,
                AutoSize = true,
                Font = new Font(SystemFonts.MessageBoxFont.FontFamily, 18, FontStyle.Bold),
                Margin = new Padding(0, 0, 0, 18)
            };
        }

        private static Label CreateLabel(string text)
        {
            return new Label
            {
                Text = text,
                Dock = DockStyle.Fill,
                AutoSize = true,
                Font = new Font(SystemFonts.MessageBoxFont.FontFamily, 11F, FontStyle.Bold),
                Margin = new Padding(0, 8, 0, 6)
            };
        }

        private static Control CreateReadonlyLine(string label, string value)
        {
            var box = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                ColumnCount = 1,
                Margin = new Padding(0, 0, 0, 14)
            };
            box.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            box.Controls.Add(CreateLabel(label));
            bool isLong = value != null && value.Length > 80;
            box.Controls.Add(new TextBox
            {
                Text = value,
                ReadOnly = true,
                Dock = DockStyle.Fill,
                Multiline = isLong,
                Height = isLong ? 72 : 30,
                ScrollBars = isLong ? ScrollBars.Horizontal : ScrollBars.None
            });
            return box;
        }

        private sealed class CurveItem
        {
            internal CurveItem(string code, string description)
            {
                Code = code;
                Description = description;
            }

            internal string Code { get; private set; }
            private string Description { get; set; }

            public override string ToString()
            {
                return Code + " - " + Description;
            }
        }
    }
}
