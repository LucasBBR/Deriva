using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace Deriva.Excel.Calendar
{
    internal static class Calendar
    {
        private static HashSet<DateTime> _holidays = new HashSet<DateTime>();
        private static List<DateTime> _duDates = new List<DateTime>();
        private static DateTime _calendarMin;
        private static DateTime _calendarMax;
        private static readonly ReaderWriterLockSlim _lock = new ReaderWriterLockSlim();
        private static bool _isLoaded = false;
        private static long _version = 0;

        internal static void SetHolidays(IEnumerable<DateTime> holidays)
        {
            _lock.EnterWriteLock();
            try
            {
                _holidays = new HashSet<DateTime>(holidays.Select(d => d.Date));
                _isLoaded = true;

                if (_holidays.Count > 0)
                {
                    _calendarMin = _holidays.Min().Date;
                    _calendarMax = _holidays.Max().Date;

                    var duList = new List<DateTime>();
                    for (var d = _calendarMin; d <= _calendarMax; d = d.AddDays(1))
                    {
                        if (d.DayOfWeek != DayOfWeek.Saturday &&
                            d.DayOfWeek != DayOfWeek.Sunday &&
                            !_holidays.Contains(d))
                            duList.Add(d);
                    }
                    _duDates = duList;
                }

                _version++;
            }
            finally
            {
                _lock.ExitWriteLock();
            }
        }

        internal static bool IsLoaded
        {
            get
            {
                _lock.EnterReadLock();
                try { return _isLoaded; }
                finally { _lock.ExitReadLock(); }
            }
        }

        internal static long Version
        {
            get
            {
                _lock.EnterReadLock();
                try { return _version; }
                finally { _lock.ExitReadLock(); }
            }
        }

        // Returns null when the holiday list has not been loaded yet.
        internal static bool? IsDU(DateTime date)
        {
            if (!IsLoaded)
                return null;

            var d = date.Date;
            if (d.DayOfWeek == DayOfWeek.Saturday) return false;
            if (d.DayOfWeek == DayOfWeek.Sunday)   return false;

            _lock.EnterReadLock();
            try
            {
                return !_holidays.Contains(d);
            }
            finally
            {
                _lock.ExitReadLock();
            }
        }

        // Returns null if not loaded. Weekends are NOT holidays — only ANBIMA dates.
        internal static bool? IsHoliday(DateTime date)
        {
            if (!IsLoaded)
                return null;

            var d = date.Date;
            _lock.EnterReadLock();
            try
            {
                return _holidays.Contains(d);
            }
            finally
            {
                _lock.ExitReadLock();
            }
        }

        // Adjusts date to the nearest business day per convention (F, MF, P, MP).
        // Returns null if not loaded. Caller must pass a valid, normalised convention string.
        internal static DateTime? AdjustToConvention(DateTime date, string convention)
        {
            if (!IsLoaded)
                return null;

            var d = date.Date;
            _lock.EnterReadLock();
            try
            {
                switch (convention)
                {
                    case "F":  return Following(d);
                    case "MF": return ModifiedFollowing(d);
                    case "P":  return Preceding(d);
                    case "MP": return ModifiedPreceding(d);
                    default:   return null;
                }
            }
            finally
            {
                _lock.ExitReadLock();
            }
        }

        // n=0  → next BD on or after date (AdjustDU "F" semantics)
        // n>0  → nth BD strictly after date (date not counted)
        // n<0  → |n|th BD strictly before date (date not counted)
        // Returns null if not loaded or result falls outside the ANBIMA data range.
        internal static DateTime? ShiftDU(DateTime date, int n)
        {
            _lock.EnterReadLock();
            try
            {
                if (!_isLoaded)
                    return null;

                return ShiftDUCore(date.Date, n);
            }
            finally
            {
                _lock.ExitReadLock();
            }
        }

        internal static DateTime?[] ShiftDUBulk(IList<DateTime> dates, int n)
        {
            if (dates == null)
                return null;

            _lock.EnterReadLock();
            try
            {
                if (!_isLoaded)
                    return null;

                var results = new DateTime?[dates.Count];
                for (int i = 0; i < dates.Count; i++)
                    results[i] = ShiftDUCore(dates[i].Date, n);
                return results;
            }
            finally
            {
                _lock.ExitReadLock();
            }
        }

        // Business days in (min(d1,d2), max(d1,d2)] — exclusive start, inclusive end.
        // DU(D, D) = 0. Order does not matter. Returns null if not loaded.
        internal static int? CountDU(DateTime d1, DateTime d2)
        {
            _lock.EnterReadLock();
            try
            {
                if (!_isLoaded)
                    return null;

                return CountDUCore(d1, d2);
            }
            finally
            {
                _lock.ExitReadLock();
            }
        }

        internal static int?[,] CountDUBulk(DateTime d1, DateTime?[,] d2s)
        {
            if (d2s == null)
                return null;

            _lock.EnterReadLock();
            try
            {
                if (!_isLoaded)
                    return null;

                int rows = d2s.GetLength(0);
                int cols = d2s.GetLength(1);
                var results = new int?[rows, cols];

                for (int r = 0; r < rows; r++)
                {
                    for (int c = 0; c < cols; c++)
                    {
                        if (d2s[r, c].HasValue)
                            results[r, c] = CountDUCore(d1, d2s[r, c].Value);
                    }
                }

                return results;
            }
            finally
            {
                _lock.ExitReadLock();
            }
        }

        // All ANBIMA holidays in [min(start,end), max(start,end)], sorted ascending.
        // Returns null if not loaded.
        internal static List<DateTime> GetHolidays(DateTime startDate, DateTime endDate)
        {
            if (!IsLoaded)
                return null;

            var min = (startDate <= endDate ? startDate : endDate).Date;
            var max = (startDate <= endDate ? endDate : startDate).Date;

            _lock.EnterReadLock();
            try
            {
                return _holidays
                    .Where(h => h >= min && h <= max)
                    .OrderBy(h => h)
                    .ToList();
            }
            finally
            {
                _lock.ExitReadLock();
            }
        }

        // --- Private helpers (called inside an already-held read lock) ---

        private static bool IsBusinessDay(DateTime d)
        {
            if (d.DayOfWeek == DayOfWeek.Saturday) return false;
            if (d.DayOfWeek == DayOfWeek.Sunday)   return false;
            return !_holidays.Contains(d);
        }

        private static DateTime? ShiftDUCore(DateTime d, int n)
        {
            if (n == 0)
                return Following(d);

            if (n > 0)
            {
                int raw = _duDates.BinarySearch(d);
                // lo = first index strictly after d
                int lo = (raw >= 0) ? raw + 1 : ~raw;
                int target = lo + n - 1;
                if (target >= _duDates.Count) return null;
                return _duDates[target];
            }
            else // n < 0
            {
                int raw = _duDates.BinarySearch(d);
                // hi = last index strictly before d
                int hi = (raw >= 0) ? raw - 1 : ~raw - 1;
                int target = hi + n + 1; // n is negative, so this subtracts |n|-1
                if (target < 0) return null;
                return _duDates[target];
            }
        }

        private static int CountDUCore(DateTime d1, DateTime d2)
        {
            var min = (d1 <= d2 ? d1 : d2).Date;
            var max = (d1 <= d2 ? d2 : d1).Date;

            // first DU strictly after min
            int rawLo = _duDates.BinarySearch(min);
            int lo = (rawLo >= 0) ? rawLo + 1 : ~rawLo;

            // last DU on or before max
            int rawHi = _duDates.BinarySearch(max);
            int hi = (rawHi >= 0) ? rawHi : ~rawHi - 1;

            return (hi >= lo) ? hi - lo + 1 : 0;
        }

        private static DateTime? Following(DateTime d)
        {
            if (IsBusinessDay(d)) return d;
            // d is not a BD: BinarySearch returns negative (not found in _duDates)
            int lo = ~_duDates.BinarySearch(d); // insertion point = first DU > d
            if (lo >= _duDates.Count) return null;
            return _duDates[lo];
        }

        private static DateTime? Preceding(DateTime d)
        {
            if (IsBusinessDay(d)) return d;
            int insertionPoint = ~_duDates.BinarySearch(d);
            int hi = insertionPoint - 1; // last DU < d
            if (hi < 0) return null;
            return _duDates[hi];
        }

        private static DateTime? ModifiedFollowing(DateTime d)
        {
            var result = Following(d);
            if (!result.HasValue) return null;
            if (result.Value.Month != d.Month)
                return Preceding(d);
            return result;
        }

        private static DateTime? ModifiedPreceding(DateTime d)
        {
            var result = Preceding(d);
            if (!result.HasValue) return null;
            if (result.Value.Month != d.Month)
                return Following(d);
            return result;
        }
    }
}
