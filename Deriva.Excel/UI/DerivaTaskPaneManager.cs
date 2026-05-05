using ExcelDna.Integration;
using System;
using System.Drawing;
using System.Windows.Forms;

namespace Deriva.Excel.UI
{
    internal static class DerivaTaskPaneManager
    {
        private static bool _visualStylesEnabled;

        internal static void ShowDashboard()
        {
            try
            {
                EnsureWinFormsInitialized();
                var dashboard = new DashboardPane();
                using (var form = CreateDialog("Deriva Dashboard", dashboard, 760, 520))
                {
                    dashboard.RefreshStatus();
                    ShowDialogWindow(form);
                }
            }
            catch (Exception ex)
            {
                ShowError("Unable to open Deriva Dashboard.", ex);
            }
        }

        internal static void ShowSettings()
        {
            try
            {
                EnsureWinFormsInitialized();
                var settings = new SettingsPane();
                using (var form = CreateDialog("Deriva Settings", settings, 940, 760))
                {
                    ShowDialogWindow(form);
                }
            }
            catch (Exception ex)
            {
                ShowError("Unable to open Deriva Settings.", ex);
            }
        }

        private static Form CreateDialog(string title, Control content, int width, int height)
        {
            var owner = new ExcelWindowOwner();
            var screen = owner.Handle == IntPtr.Zero
                ? Screen.PrimaryScreen
                : Screen.FromHandle(owner.Handle);
            var area = screen.WorkingArea;

            int dialogWidth = Math.Max(width, (int)(area.Width * 0.42));
            int dialogHeight = Math.Max(height, (int)(area.Height * 0.68));
            dialogWidth = Math.Min(dialogWidth, Math.Max(640, area.Width - 96));
            dialogHeight = Math.Min(dialogHeight, Math.Max(480, area.Height - 96));

            var form = new Form
            {
                Text = title,
                StartPosition = FormStartPosition.CenterParent,
                Size = new Size(dialogWidth, dialogHeight),
                MinimumSize = new Size(
                    Math.Min(width, Math.Max(640, area.Width - 160)),
                    Math.Min(height, Math.Max(480, area.Height - 160))),
                AutoScaleMode = AutoScaleMode.Dpi,
                Font = new Font("Segoe UI", area.Width >= 2500 ? 11.5f : 10.5f),
                ShowIcon = false,
                ShowInTaskbar = false,
                MaximizeBox = true
            };

            content.Font = form.Font;
            content.Dock = DockStyle.Fill;
            form.Controls.Add(content);
            return form;
        }

        private static void ShowDialogWindow(Form form)
        {
            form.ShowDialog(new ExcelWindowOwner());
        }

        private static void EnsureWinFormsInitialized()
        {
            if (_visualStylesEnabled)
                return;

            try
            {
                Application.EnableVisualStyles();
            }
            catch
            {
                // Excel may already have initialized WinForms; this is only cosmetic.
            }

            _visualStylesEnabled = true;
        }

        private sealed class ExcelWindowOwner : IWin32Window
        {
            public IntPtr Handle
            {
                get { return ExcelDnaUtil.WindowHandle; }
            }
        }

        private static void ShowError(string message, Exception ex)
        {
            MessageBox.Show(
                message + Environment.NewLine + ex.Message,
                "Deriva",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
    }
}
