using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Deriva.Excel.Calendar;
using Deriva.Excel.Diagnostics;
using Deriva.Excel.ETTJ;

namespace Deriva.Excel
{
    internal static class MarketDataRefreshService
    {
        internal static async Task RefreshAllAsync(bool forceEttjDownload)
        {
            DerivaLog.Info("RefreshAllAsync started. forceEttjDownload=" + forceEttjDownload);
            if (forceEttjDownload)
                EttjMemoryCache.Clear();

            var settings = EttjSettings.Load();
            await RefreshHolidaysAsync(settings).ConfigureAwait(false);
            await RefreshInitialEttjAsync(settings, forceEttjDownload).ConfigureAwait(false);
            DerivaLog.Info("RefreshAllAsync finished.");
        }

        internal static async Task RefreshHolidaysAsync(EttjSettings settings)
        {
            DerivaLog.Info("Holiday refresh started. Url=" + settings.HolidayUrl);
            List<DateTime> freshDates = null;
            Exception holidayError = null;

            try
            {
                freshDates = await HolidayFetcher.FetchAsync(settings.HolidayUrl).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                holidayError = ex;
                DerivaLog.Error("Holiday fetch failed.", ex);
            }

            if (freshDates != null)
            {
                HolidayCache.Save(freshDates);
                Calendar.Calendar.SetHolidays(freshDates);
                EttjStatusStore.RecordSuccess(
                    "Holiday",
                    "Fetched " + freshDates.Count + " holidays from " + settings.HolidayUrl);
                DerivaLog.Info("Holiday refresh succeeded. Count=" + freshDates.Count);
                return;
            }

            var cached = HolidayCache.TryLoad();
            if (cached.HasValue)
            {
                var (dates, isStale) = cached.Value;
                Calendar.Calendar.SetHolidays(dates);
                EttjStatusStore.RecordError(
                    "Holiday",
                    "Using cached holidays. " +
                    (holidayError == null ? "Fresh fetch failed." : holidayError.Message));
                DerivaLog.Warn("Holiday refresh using cache. Count=" + dates.Count + ", isStale=" + isStale);
                return;
            }

            EttjStatusStore.RecordError(
                "Holiday",
                holidayError == null
                    ? "Failed to load ANBIMA holidays and no cache was found."
                    : holidayError.Message);
            DerivaLog.Error("Holiday refresh failed and no cache was found.", holidayError);
        }

        internal static async Task RefreshInitialEttjAsync(
            EttjSettings settings,
            bool forceDownload)
        {
            DerivaLog.Info("ETTJ refresh started. Curves=" + settings.InitCurves + ", forceDownload=" + forceDownload);
            try
            {
                var refDate = PreviousBusinessDay(DateTime.Today);
                if (!refDate.HasValue)
                {
                    var detail = "Holiday calendar is not loaded; previous business day could not be resolved.";
                    EttjStatusStore.RecordError("ETTJ", detail);
                    DerivaLog.Warn("ETTJ refresh skipped. " + detail);
                    return;
                }

                var curves = EttjService.ParseCurveCsv(settings.InitCurves);
                var service = new EttjService(settings);
                var records = await service.GetEttjAsync(refDate.Value, curves, !forceDownload)
                    .ConfigureAwait(false);

                EttjStatusStore.RecordSuccess(
                    "ETTJ",
                    "Fetched " + records.Count + " rows for " +
                    refDate.Value.ToString("yyyy-MM-dd") +
                    " (" + string.Join(",", curves.ToArray()) + ")");
                DerivaLog.Info(
                    "ETTJ refresh succeeded. RefDate=" + refDate.Value.ToString("yyyy-MM-dd") +
                    ", rows=" + records.Count +
                    ", curves=" + string.Join(",", curves.ToArray()));
            }
            catch (Exception ex)
            {
                EttjStatusStore.RecordError("ETTJ", ex.Message);
                DerivaLog.Error("ETTJ refresh failed.", ex);
            }
        }

        internal static DateTime? PreviousBusinessDay(DateTime date)
        {
            var cursor = date.Date.AddDays(-1);
            for (int i = 0; i < 14; i++)
            {
                var isBusinessDay = Calendar.Calendar.IsDU(cursor);
                if (!isBusinessDay.HasValue)
                    return null;
                if (isBusinessDay.Value)
                    return cursor;
                cursor = cursor.AddDays(-1);
            }

            return null;
        }
    }
}
