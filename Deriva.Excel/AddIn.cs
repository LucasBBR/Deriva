using System;
using System.Threading.Tasks;
using Deriva.Excel.Diagnostics;
using ExcelDna.Integration;
using ExcelDna.Logging;

namespace Deriva.Excel
{
    public class AddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            DerivaLog.Info("AutoOpen started.");
            Calendar.HolidayFetcher.ConfigureTls();
            ExcelAsyncUtil.QueueAsMacro(() => { _ = LoadHolidaysAsync(); });
        }

        public void AutoClose()
        {
            DerivaLog.Info("AutoClose.");
        }

        private static async Task LoadHolidaysAsync()
        {
            try
            {
                await MarketDataRefreshService.RefreshAllAsync(false).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                DerivaLog.Error("AutoOpen refresh failed.", ex);
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    LogDisplay.RecordLine(
                        "Deriva: refresh failed during startup. " + ex.Message);
                    LogDisplay.Show();
                });
            }
        }

        private static void ShowStatusBarWarning(string message)
        {
            XlCall.Excel(XlCall.xlcMessage, true, message);
            Task.Delay(8000).ContinueWith(_ =>
                ExcelAsyncUtil.QueueAsMacro(() =>
                    XlCall.Excel(XlCall.xlcMessage, false, "")));
        }
    }
}
