using System;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;
using Deriva.Excel.Diagnostics;
using Deriva.Excel.ETTJ;

namespace Deriva.Excel.UI
{
    internal sealed class DashboardPane : UserControl
    {
        private readonly Label _holidayTime;
        private readonly Label _holidayStatus;
        private readonly Label _holidayDetail;
        private readonly Label _ettjTime;
        private readonly Label _ettjStatus;
        private readonly Label _ettjDetail;
        private readonly Button _refreshButton;

        internal DashboardPane()
        {
            Dock = DockStyle.Fill;
            BackColor = Color.White;
            AutoScaleMode = AutoScaleMode.Dpi;

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(24),
                ColumnCount = 1,
                RowCount = 5,
                AutoScroll = true
            };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            Controls.Add(root);

            var title = new Label
            {
                Text = "Updates",
                Dock = DockStyle.Fill,
                AutoSize = true,
                Font = new Font(Font.FontFamily, 18, FontStyle.Bold),
                Margin = new Padding(0, 0, 0, 18)
            };
            root.Controls.Add(title);

            _holidayTime = new Label();
            _holidayStatus = new Label();
            _holidayDetail = new Label();
            root.Controls.Add(CreateStatusPanel("Holiday", _holidayTime, _holidayStatus, _holidayDetail));

            _ettjTime = new Label();
            _ettjStatus = new Label();
            _ettjDetail = new Label();
            root.Controls.Add(CreateStatusPanel("ETTJ", _ettjTime, _ettjStatus, _ettjDetail));

            _refreshButton = new Button
            {
                Text = "Refresh",
                AutoSize = true,
                MinimumSize = new Size(120, 38),
                Margin = new Padding(0, 18, 0, 0),
                Padding = new Padding(16, 6, 16, 6)
            };
            _refreshButton.Click += async (sender, args) => await RefreshNowAsync();
            root.Controls.Add(_refreshButton);

            RefreshStatus();
        }

        internal void RefreshStatus()
        {
            var snapshot = EttjStatusStore.Load();
            ApplyStatus(snapshot.Holiday, _holidayTime, _holidayStatus, _holidayDetail);
            ApplyStatus(snapshot.ETTJ, _ettjTime, _ettjStatus, _ettjDetail);
        }

        private async Task RefreshNowAsync()
        {
            DerivaLog.Info("Dashboard refresh clicked.");
            SetLoading();
            try
            {
                await MarketDataRefreshService.RefreshAllAsync(true).ConfigureAwait(true);
            }
            catch (Exception ex)
            {
                DerivaLog.Error("Dashboard refresh failed.", ex);
                MessageBox.Show(
                    "Refresh failed." + Environment.NewLine + ex.Message,
                    "Deriva Dashboard",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                RefreshStatus();
                _refreshButton.Enabled = true;
                _refreshButton.Text = "Refresh";
                DerivaLog.Info("Dashboard refresh UI updated.");
            }
        }

        private void SetLoading()
        {
            _refreshButton.Enabled = false;
            _refreshButton.Text = "Refreshing...";
            _holidayStatus.Text = "Status: Loading...";
            _holidayStatus.ForeColor = Color.DimGray;
            _holidayDetail.Text = "Detail: Fetching holidays...";
            _ettjStatus.Text = "Status: Loading...";
            _ettjStatus.ForeColor = Color.DimGray;
            _ettjDetail.Text = "Detail: Fetching ETTJ...";
        }

        private static Control CreateStatusPanel(
            string name,
            Label timeLabel,
            Label statusLabel,
            Label detailLabel)
        {
            var panel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                ColumnCount = 1,
                Margin = new Padding(0, 0, 0, 18),
                Padding = new Padding(16),
                BackColor = Color.FromArgb(248, 249, 250)
            };
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

            panel.Controls.Add(new Label
            {
                Text = name,
                Dock = DockStyle.Fill,
                AutoSize = true,
                Font = new Font(SystemFonts.MessageBoxFont.FontFamily, 12, FontStyle.Bold),
                Margin = new Padding(0, 0, 0, 10)
            });

            ConfigureValueLabel(timeLabel);
            ConfigureValueLabel(statusLabel);
            ConfigureValueLabel(detailLabel);

            panel.Controls.Add(timeLabel);
            panel.Controls.Add(statusLabel);
            panel.Controls.Add(detailLabel);
            return panel;
        }

        private static void ConfigureValueLabel(Label label)
        {
            label.Dock = DockStyle.Fill;
            label.AutoSize = true;
            label.MaximumSize = new Size(0, 0);
            label.Font = new Font(SystemFonts.MessageBoxFont.FontFamily, 11F);
            label.Margin = new Padding(0, 4, 0, 4);
        }

        private static void ApplyStatus(
            StatusEntry entry,
            Label timeLabel,
            Label statusLabel,
            Label detailLabel)
        {
            timeLabel.Text = "Last update: " +
                (entry.LastUpdateLocal.HasValue
                    ? entry.LastUpdateLocal.Value.ToString("yyyy-MM-dd HH:mm:ss")
                    : "-");
            statusLabel.Text = "Status: " + entry.Status;
            detailLabel.Text = string.IsNullOrWhiteSpace(entry.Detail)
                ? "Detail: -"
                : "Detail: " + entry.Detail;

            if (entry.Status == "Success")
                statusLabel.ForeColor = Color.FromArgb(24, 128, 74);
            else if (entry.Status == "Error")
                statusLabel.ForeColor = Color.FromArgb(180, 50, 50);
            else
                statusLabel.ForeColor = Color.DimGray;
        }
    }
}
