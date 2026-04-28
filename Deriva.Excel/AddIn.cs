using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Deriva.Excel.Calendar;
using ExcelDna.Integration;
using ExcelDna.Logging;

namespace Deriva.Excel
{
    public class AddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            HolidayFetcher.ConfigureTls();
            ExcelAsyncUtil.QueueAsMacro(() => { _ = LoadHolidaysAsync(); });
        }

        public void AutoClose() { }

        private static async Task LoadHolidaysAsync()
        {
            List<DateTime> freshDates = null;
            try
            {
                freshDates = await HolidayFetcher.FetchAsync().ConfigureAwait(false);
            }
            catch { }

            if (freshDates != null)
            {
                HolidayCache.Save(freshDates);
                ExcelAsyncUtil.QueueAsMacro(() => Calendar.Calendar.SetHolidays(freshDates));
                return;
            }

            var cached = HolidayCache.TryLoad();
            if (cached.HasValue)
            {
                var (dates, isStale) = cached.Value;
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    Calendar.Calendar.SetHolidays(dates);
                    if (isStale)
                        ShowStatusBarWarning(
                            "Deriva: Não foi possível atualizar feriados ANBIMA. Usando cache local.");
                });
            }
            else
            {
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    LogDisplay.RecordLine(
                        "Deriva: Falha ao carregar feriados ANBIMA e nenhum cache encontrado. " +
                        "A função =IsDU() não estará disponível.");
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
