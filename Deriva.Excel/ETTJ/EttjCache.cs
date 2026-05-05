using System;
using System.Collections.Generic;
using System.IO;
using Deriva.Excel.Diagnostics;
using Newtonsoft.Json;

namespace Deriva.Excel.ETTJ
{
    internal sealed class EttjCache
    {
        private readonly string _cacheDir;

        internal EttjCache(EttjSettings settings)
        {
            var normalized = (settings ?? EttjSettings.Load()).Normalized();
            _cacheDir = Path.Combine(normalized.EttjCacheDir, "TS");
        }

        internal bool IsValid(DateTime refDate)
        {
            var path = GetPath(refDate);
            DerivaLog.Info("ETTJ cache validity check. Path=" + path);
            if (!File.Exists(path))
            {
                DerivaLog.Info("ETTJ cache miss.");
                return false;
            }

            var date = refDate.Date;
            var today = DateTime.Today;
            if (date >= today)
            {
                DerivaLog.Info("ETTJ cache invalid because date is today/future.");
                return false;
            }

            if ((today - date).TotalDays > 5)
            {
                DerivaLog.Info("ETTJ cache valid as immutable old data.");
                return true;
            }

            try
            {
                var age = DateTime.Now - File.GetLastWriteTime(path);
                DerivaLog.Info("ETTJ cache age hours=" + age.TotalHours);
                return age.TotalHours < 4.0;
            }
            catch
            {
                return false;
            }
        }

        internal List<EttjRecord> TryLoad(DateTime refDate)
        {
            var path = GetPath(refDate);
            if (!File.Exists(path))
                return null;

            try
            {
                var json = File.ReadAllText(path);
                var records = JsonConvert.DeserializeObject<List<EttjRecord>>(json);
                DerivaLog.Info("ETTJ cache loaded rows=" + (records == null ? 0 : records.Count) + ". Path=" + path);
                return records;
            }
            catch (Exception ex)
            {
                DerivaLog.Error("ETTJ cache load failed. Path=" + path, ex);
                return null;
            }
        }

        internal void Save(DateTime refDate, List<EttjRecord> records)
        {
            if (records == null || records.Count == 0)
                return;

            Directory.CreateDirectory(_cacheDir);
            File.WriteAllText(
                GetPath(refDate),
                JsonConvert.SerializeObject(records, Formatting.Indented));
            DerivaLog.Info("ETTJ cache saved rows=" + records.Count + ". Path=" + GetPath(refDate));
        }

        private string GetPath(DateTime refDate)
        {
            return Path.Combine(_cacheDir, refDate.Date.ToString("yyyy-MM-dd") + ".json");
        }
    }
}
