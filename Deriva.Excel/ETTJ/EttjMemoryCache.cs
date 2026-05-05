using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Deriva.Excel.Diagnostics;
using Deriva.Excel.Interpolation;

namespace Deriva.Excel.ETTJ
{
    internal static class EttjMemoryCache
    {
        private static readonly ConcurrentDictionary<string, Lazy<Task<List<EttjRecord>>>> DayCache =
            new ConcurrentDictionary<string, Lazy<Task<List<EttjRecord>>>>(StringComparer.Ordinal);

        private static readonly ConcurrentDictionary<string, Lazy<Task<PreparedEttjCurve>>> CurveCache =
            new ConcurrentDictionary<string, Lazy<Task<PreparedEttjCurve>>>(StringComparer.Ordinal);

        internal static async Task<PreparedEttjCurve> GetPreparedCurveAsync(
            DateTime referenceDate,
            string curve,
            InterpolationMethod method)
        {
            var settings = EttjSettings.Load().Normalized();
            var date = referenceDate.Date;
            var normalizedCurve = curve.Trim().ToUpperInvariant();
            var cacheKey = BuildCurveCacheKey(settings, date, normalizedCurve, method);

            if (ShouldBypassMemoryCache(date))
                return await BuildPreparedCurveAsync(settings, date, normalizedCurve, method).ConfigureAwait(false);

            var lazy = CurveCache.GetOrAdd(
                cacheKey,
                key =>
                {
                    DerivaLog.Info(
                        "ETTJ memory curve cache miss. RefDate=" + date.ToString("yyyy-MM-dd") +
                        ", curve=" + normalizedCurve +
                        ", method=" + method);
                    return new Lazy<Task<PreparedEttjCurve>>(
                        () => BuildPreparedCurveAsync(settings, date, normalizedCurve, method),
                        LazyThreadSafetyMode.ExecutionAndPublication);
                });

            try
            {
                return await lazy.Value.ConfigureAwait(false);
            }
            catch
            {
                Lazy<Task<PreparedEttjCurve>> ignored;
                CurveCache.TryRemove(cacheKey, out ignored);
                throw;
            }
        }

        internal static void Clear()
        {
            DayCache.Clear();
            CurveCache.Clear();
            DerivaLog.Info("ETTJ memory caches cleared.");
        }

        private static async Task<PreparedEttjCurve> BuildPreparedCurveAsync(
            EttjSettings settings,
            DateTime referenceDate,
            string curve,
            InterpolationMethod method)
        {
            var records = await GetDayAsync(settings, referenceDate).ConfigureAwait(false);

            var points = records
                .Where(r => string.Equals(r.Curva, curve, StringComparison.OrdinalIgnoreCase))
                .Where(r => r.DiasUteis > 0)
                .GroupBy(r => r.DiasUteis)
                .Select(g => g.OrderBy(r => r.DiasCorridos).First())
                .OrderBy(r => r.DiasUteis)
                .ToList();

            if (points.Count < 2)
                throw new EttjNoDataException("Requested ETTJ curve has fewer than two valid points.");

            var xs = points.Select(r => (double)r.DiasUteis).ToArray();
            var ys = points.Select(r => r.Taxa).ToArray();

            PreparedInterpolator interpolator;
            var result = PreparedInterpolator.TryCreate(xs, ys, method, out interpolator);
            if (result.Error != InterpolationError.None)
                throw new EttjException("Unable to prepare ETTJ interpolation curve: " + result.Error);

            DerivaLog.Info(
                "ETTJ prepared curve built. RefDate=" + referenceDate.ToString("yyyy-MM-dd") +
                ", curve=" + curve +
                ", points=" + points.Count +
                ", minDU=" + interpolator.MinX.ToString(CultureInfo.InvariantCulture) +
                ", maxDU=" + interpolator.MaxX.ToString(CultureInfo.InvariantCulture));

            return new PreparedEttjCurve(curve, interpolator);
        }

        private static async Task<List<EttjRecord>> GetDayAsync(
            EttjSettings settings,
            DateTime referenceDate)
        {
            if (ShouldBypassMemoryCache(referenceDate))
                return await LoadFullDayAsync(settings, referenceDate).ConfigureAwait(false);

            var cacheKey = BuildDayCacheKey(settings, referenceDate);
            var lazy = DayCache.GetOrAdd(
                cacheKey,
                key =>
                {
                    DerivaLog.Info("ETTJ memory day cache miss. RefDate=" + referenceDate.ToString("yyyy-MM-dd"));
                    return new Lazy<Task<List<EttjRecord>>>(
                        () => LoadFullDayAsync(settings, referenceDate),
                        LazyThreadSafetyMode.ExecutionAndPublication);
                });

            try
            {
                return await lazy.Value.ConfigureAwait(false);
            }
            catch
            {
                Lazy<Task<List<EttjRecord>>> ignored;
                DayCache.TryRemove(cacheKey, out ignored);
                throw;
            }
        }

        private static async Task<List<EttjRecord>> LoadFullDayAsync(
            EttjSettings settings,
            DateTime referenceDate)
        {
            var service = new EttjService(settings);
            return await service.GetEttjAsync(
                    referenceDate,
                    EttjCurveCatalog.AvailableCurves.Keys,
                    true)
                .ConfigureAwait(false);
        }

        private static bool ShouldBypassMemoryCache(DateTime referenceDate)
        {
            return referenceDate.Date >= DateTime.Today;
        }

        private static string BuildCurveCacheKey(
            EttjSettings settings,
            DateTime referenceDate,
            string curve,
            InterpolationMethod method)
        {
            return BuildDayCacheKey(settings, referenceDate) +
                   "|" + curve +
                   "|" + method;
        }

        private static string BuildDayCacheKey(EttjSettings settings, DateTime referenceDate)
        {
            return (settings.EttjCacheDir ?? string.Empty).Trim().ToUpperInvariant() +
                   "|" + referenceDate.ToString("yyyyMMdd", CultureInfo.InvariantCulture) +
                   "|" + GetCacheBucket(referenceDate);
        }

        private static string GetCacheBucket(DateTime referenceDate)
        {
            var date = referenceDate.Date;
            var today = DateTime.Today;
            if ((today - date).TotalDays > 5)
                return "immutable";

            long fourHourTicks = TimeSpan.FromHours(4).Ticks;
            return (DateTime.Now.Ticks / fourHourTicks).ToString(CultureInfo.InvariantCulture);
        }
    }
}
