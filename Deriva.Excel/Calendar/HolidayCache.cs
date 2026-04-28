using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Newtonsoft.Json;

namespace Deriva.Excel.Calendar
{
    internal static class HolidayCache
    {
        private static readonly string CacheDir =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Deriva");

        private static readonly string CacheFile =
            Path.Combine(CacheDir, "holidays.json");

        private static readonly TimeSpan Ttl = TimeSpan.FromDays(1);

        internal static void Save(List<DateTime> holidays)
        {
            Directory.CreateDirectory(CacheDir);
            var payload = new CachePayload
            {
                LastUpdated = DateTime.UtcNow,
                Holidays = holidays.Select(d => d.ToString("yyyy-MM-dd")).ToList()
            };
            File.WriteAllText(CacheFile, JsonConvert.SerializeObject(payload, Formatting.Indented));
        }

        internal static (List<DateTime> Dates, bool IsStale)? TryLoad()
        {
            if (!File.Exists(CacheFile))
                return null;

            try
            {
                var json = File.ReadAllText(CacheFile);
                var payload = JsonConvert.DeserializeObject<CachePayload>(json);
                if (payload?.Holidays == null)
                    return null;

                var dates = payload.Holidays
                    .Select(s => DateTime.ParseExact(
                                s, "yyyy-MM-dd",
                                System.Globalization.CultureInfo.InvariantCulture))
                    .ToList();

                bool isStale = (DateTime.UtcNow - payload.LastUpdated) > Ttl;
                return (dates, isStale);
            }
            catch
            {
                return null;
            }
        }

        private class CachePayload
        {
            public DateTime LastUpdated { get; set; }
            public List<string> Holidays { get; set; }
        }
    }
}
