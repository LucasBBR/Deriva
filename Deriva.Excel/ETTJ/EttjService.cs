using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Deriva.Excel.Diagnostics;

namespace Deriva.Excel.ETTJ
{
    internal sealed class EttjService
    {
        private readonly EttjClient _client;
        private readonly EttjCache _cache;

        internal EttjService(EttjSettings settings)
        {
            _client = new EttjClient();
            _cache = new EttjCache(settings);
        }

        internal async Task<List<EttjRecord>> GetEttjAsync(
            DateTime refDate,
            IEnumerable<string> curves,
            bool useCache)
        {
            var date = refDate.Date;
            DerivaLog.Info(
                "GetEttjAsync start. RefDate=" + date.ToString("yyyy-MM-dd") +
                ", curves=" + string.Join(",", curves == null ? new string[0] : curves.ToArray()) +
                ", useCache=" + useCache);
            if (date.DayOfWeek == DayOfWeek.Saturday || date.DayOfWeek == DayOfWeek.Sunday)
                throw new EttjNoDataException("No B3 TaxaSwap data for weekends.");

            if (date > DateTime.Today)
                throw new EttjNoDataException("No B3 TaxaSwap data for future dates.");

            var requestedCurves = EttjCurveCatalog.Normalize(curves);
            DerivaLog.Info("GetEttjAsync requested curves normalized=" + string.Join(",", requestedCurves.ToArray()));

            if (useCache && _cache.IsValid(date))
            {
                var cached = _cache.TryLoad(date);
                if (cached != null)
                {
                    var filtered = Filter(cached, requestedCurves);
                    if (filtered.Count > 0)
                    {
                        DerivaLog.Info("GetEttjAsync returning cache rows=" + filtered.Count);
                        return filtered;
                    }
                }
            }

            var downloadCurves = useCache
                ? EttjCurveCatalog.AvailableCurves.Keys
                : requestedCurves;

            var fullDayRecords = await _client.FetchAsync(date, downloadCurves).ConfigureAwait(false);
            DerivaLog.Info("GetEttjAsync downloaded rows=" + fullDayRecords.Count);
            if (fullDayRecords.Count == 0)
                throw new EttjNoDataException("No B3 TaxaSwap records found for the requested date.");

            if (useCache)
                _cache.Save(date, fullDayRecords);

            var result = Filter(fullDayRecords, requestedCurves);
            if (result.Count == 0)
                throw new EttjNoDataException("Requested ETTJ curve was not found for the requested date.");

            DerivaLog.Info("GetEttjAsync returning rows=" + result.Count);
            return result;
        }

        internal async Task<List<EttjRecord>> GetHistoricoAsync(
            DateTime startDate,
            DateTime endDate,
            IEnumerable<string> curves,
            bool useCache,
            bool ignoreErrors)
        {
            var start = startDate.Date;
            var end = endDate.Date;
            if (start > end)
                throw new EttjException("Start date is after end date.");

            var records = new List<EttjRecord>();
            for (var date = start; date <= end; date = date.AddDays(1))
            {
                if (date.DayOfWeek == DayOfWeek.Saturday || date.DayOfWeek == DayOfWeek.Sunday)
                    continue;

                try
                {
                    var daily = await GetEttjAsync(date, curves, useCache).ConfigureAwait(false);
                    records.AddRange(daily);
                }
                catch (EttjNoDataException)
                {
                    if (!ignoreErrors)
                        throw;
                }
                catch (EttjException)
                {
                    if (!ignoreErrors)
                        throw;
                }
            }

            if (records.Count == 0)
                throw new EttjNoDataException("No ETTJ data was returned for the requested interval.");

            return records
                .OrderBy(r => r.RefDate)
                .ThenBy(r => r.Curva, StringComparer.Ordinal)
                .ThenBy(r => r.DiasCorridos)
                .ToList();
        }

        internal static List<string> ParseCurveCsv(string curveText)
        {
            if (string.IsNullOrWhiteSpace(curveText))
                return EttjCurveCatalog.Normalize(null);

            var parts = curveText
                .Split(new[] { ',', ';', '|', ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(s => s.Trim());

            return EttjCurveCatalog.Normalize(parts);
        }

        private static List<EttjRecord> Filter(List<EttjRecord> records, List<string> curves)
        {
            var set = new HashSet<string>(curves, StringComparer.OrdinalIgnoreCase);
            return records
                .Where(r => set.Contains(r.Curva))
                .OrderBy(r => r.Curva, StringComparer.Ordinal)
                .ThenBy(r => r.DiasCorridos)
                .ToList();
        }
    }
}
