using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;
using Deriva.Excel.Diagnostics;
using Deriva.Excel.ETTJ;
using Deriva.Excel.Interpolation;
using ExcelDna.Integration;

namespace Deriva.Excel.Functions
{
    public static class EttjFunctions
    {
        [ExcelFunction(
            Name = "GetETTJ",
            Description = "Fetches B3 ETTJ TaxaSwap curves for one date and returns a spill table.",
            Category = "Deriva - ETTJ",
            IsThreadSafe = false)]
        public static object GetETTJ(
            [ExcelArgument(Name = "data", Description = "Reference date, as an Excel date or dd/mm/yyyy.")]
            object dateArg,
            [ExcelArgument(Name = "curva", Description = "Optional curve code, comma-separated codes, range, or TODOS. Default = PRE.")]
            object curveArg = null,
            [ExcelArgument(Name = "cache", Description = "Optional TRUE/FALSE. Default = TRUE.")]
            object cacheArg = null)
        {
            DateTime refDate;
            List<string> curves;
            bool useCache;

            if (!TryParseDate(dateArg, out refDate) ||
                !TryParseCurves(curveArg, out curves) ||
                !TryParseBool(cacheArg, true, out useCache))
                return ExcelError.ExcelErrorValue;

            var key = BuildKey(refDate, refDate, curves, useCache, true);
            return ExcelAsyncUtil.RunTask("GetETTJ", key, async () =>
            {
                try
                {
                    var service = new EttjService(EttjSettings.Load());
                    var records = await service.GetEttjAsync(refDate, curves, useCache).ConfigureAwait(false);
                    EttjStatusStore.RecordSuccess(
                        "ETTJ",
                        "Fetched " + records.Count + " rows for " + refDate.ToString("yyyy-MM-dd"));
                    return (object)ToSpillTable(records);
                }
                catch (Exception ex)
                {
                    EttjStatusStore.RecordError("ETTJ", ex.Message);
                    return MapExceptionToExcelError(ex);
                }
            });
        }

        [ExcelFunction(
            Name = "GetETTJHistorico",
            Description = "Fetches B3 ETTJ TaxaSwap curves for a date interval and returns a long spill table.",
            Category = "Deriva - ETTJ",
            IsThreadSafe = false)]
        public static object GetETTJHistorico(
            [ExcelArgument(Name = "data_ini", Description = "Start date, as an Excel date or dd/mm/yyyy.")]
            object startDateArg,
            [ExcelArgument(Name = "data_fim", Description = "End date, as an Excel date or dd/mm/yyyy.")]
            object endDateArg,
            [ExcelArgument(Name = "curva", Description = "Optional curve code, comma-separated codes, range, or TODOS. Default = PRE.")]
            object curveArg = null,
            [ExcelArgument(Name = "cache", Description = "Optional TRUE/FALSE. Default = TRUE.")]
            object cacheArg = null,
            [ExcelArgument(Name = "ignorar_erros", Description = "Optional TRUE/FALSE. Default = TRUE.")]
            object ignoreErrorsArg = null)
        {
            DateTime startDate;
            DateTime endDate;
            List<string> curves;
            bool useCache;
            bool ignoreErrors;

            if (!TryParseDate(startDateArg, out startDate) ||
                !TryParseDate(endDateArg, out endDate) ||
                !TryParseCurves(curveArg, out curves) ||
                !TryParseBool(cacheArg, true, out useCache) ||
                !TryParseBool(ignoreErrorsArg, true, out ignoreErrors))
                return ExcelError.ExcelErrorValue;

            var key = BuildKey(startDate, endDate, curves, useCache, ignoreErrors);
            return ExcelAsyncUtil.RunTask("GetETTJHistorico", key, async () =>
            {
                try
                {
                    var service = new EttjService(EttjSettings.Load());
                    var records = await service.GetHistoricoAsync(
                            startDate,
                            endDate,
                            curves,
                            useCache,
                            ignoreErrors)
                        .ConfigureAwait(false);

                    EttjStatusStore.RecordSuccess(
                        "ETTJ",
                        "Fetched " + records.Count + " historical rows from " +
                        startDate.ToString("yyyy-MM-dd") + " to " +
                        endDate.ToString("yyyy-MM-dd"));
                    return (object)ToSpillTable(records);
                }
                catch (Exception ex)
                {
                    EttjStatusStore.RecordError("ETTJ", ex.Message);
                    return MapExceptionToExcelError(ex);
                }
            });
        }

        [ExcelFunction(
            Name = "GetCurve",
            Description = "Fetches one B3 ETTJ curve and returns the interpolated rate for a maturity date.",
            Category = "Deriva - ETTJ",
            IsThreadSafe = false)]
        public static object GetCurve(
            [ExcelArgument(Name = "CURVE", Description = "Curve code, such as PRE, DIC, DOC, DCL.")]
            object curveArg,
            [ExcelArgument(Name = "ReferenceDate", Description = "ETTJ reference date, as an Excel date or dd/mm/yyyy.")]
            object referenceDateArg,
            [ExcelArgument(Name = "EndDate", Description = "Maturity date used to compute DU from ReferenceDate.")]
            object endDateArg)
        {
            DateTime referenceDate;
            DateTime endDate;
            string curve;
            EndDateInput endDateInput;

            if (!TryParseSingleCurve(curveArg, out curve) ||
                !TryParseDate(referenceDateArg, out referenceDate) ||
                !TryParseEndDateInput(endDateArg, out endDateInput))
                return ExcelError.ExcelErrorValue;

            if (endDateInput.IsRange)
            {
                var rangeKey = BuildGetCurveKey(curve, referenceDate, endDateInput);
                return ExcelAsyncUtil.RunTask("GetCurve", rangeKey, async () =>
                {
                    try
                    {
                        return await CalculateGetCurveRangeAsync(curve, referenceDate, endDateInput)
                            .ConfigureAwait(false);
                    }
                    catch (Exception ex)
                    {
                        DerivaLog.Error("GetCurve range failed.", ex);
                        EttjStatusStore.RecordError("ETTJ", ex.Message);
                        return (object)endDateInput.CreateSourceErrorMatrix(
                            MapExceptionToExcelError(ex),
                            referenceDate);
                    }
                });
            }

            endDate = endDateInput.ScalarDate;

            var targetBusinessDays = GetTargetBusinessDays(referenceDate, endDate);
            if (!targetBusinessDays.HasValue)
                return endDate.Date <= referenceDate.Date
                    ? (object)ExcelError.ExcelErrorNum
                    : ExcelError.ExcelErrorNA;

            if (targetBusinessDays.Value <= 0)
                return ExcelError.ExcelErrorNum;

            var key = curve + "|" +
                      referenceDate.ToString("yyyyMMdd", CultureInfo.InvariantCulture) +
                      "|" + endDate.ToString("yyyyMMdd", CultureInfo.InvariantCulture) +
                      "|" + targetBusinessDays.Value.ToString(CultureInfo.InvariantCulture);

            return ExcelAsyncUtil.RunTask("GetCurve", key, async () =>
            {
                try
                {
                    return await CalculateGetCurveScalarAsync(curve, referenceDate, targetBusinessDays.Value)
                        .ConfigureAwait(false);
                }
                catch (Exception ex)
                {
                    DerivaLog.Error("GetCurve failed.", ex);
                    EttjStatusStore.RecordError("ETTJ", ex.Message);
                    return MapExceptionToExcelError(ex);
                }
            });
        }

        private static async Task<object> CalculateGetCurveScalarAsync(
            string curve,
            DateTime referenceDate,
            int targetBusinessDays)
        {
            var preparedCurve = await EttjMemoryCache.GetPreparedCurveAsync(
                    referenceDate,
                    curve,
                    InterpolationMethod.CubicSpline)
                .ConfigureAwait(false);

            return MapInterpolationResult(preparedCurve.Interpolate(targetBusinessDays));
        }

        private static async Task<object> CalculateGetCurveRangeAsync(
            string curve,
            DateTime referenceDate,
            EndDateInput endDates)
        {
            DerivaLog.Info(
                "GetCurve range started. Curve=" + curve +
                ", referenceDate=" + referenceDate.ToString("yyyy-MM-dd") +
                ", rows=" + endDates.Rows +
                ", cols=" + endDates.Columns);

            var counts = Calendar.Calendar.CountDUBulk(referenceDate, endDates.Dates);
            if (counts == null)
                return endDates.CreateSourceErrorMatrix(ExcelError.ExcelErrorNA, referenceDate);

            var output = new object[endDates.Rows, endDates.Columns];
            var targets = new int?[endDates.Rows, endDates.Columns];
            int candidateCount = 0;
            int localErrorCount = 0;

            for (int r = 0; r < endDates.Rows; r++)
            {
                for (int c = 0; c < endDates.Columns; c++)
                {
                    if (!endDates.IsValid(r, c))
                    {
                        output[r, c] = ExcelError.ExcelErrorValue;
                        localErrorCount++;
                        continue;
                    }

                    var endDate = endDates.Dates[r, c].Value;
                    if (endDate.Date <= referenceDate.Date)
                    {
                        output[r, c] = ExcelError.ExcelErrorNum;
                        localErrorCount++;
                        continue;
                    }

                    if (!counts[r, c].HasValue)
                    {
                        output[r, c] = ExcelError.ExcelErrorNA;
                        localErrorCount++;
                        continue;
                    }

                    if (counts[r, c].Value <= 0)
                    {
                        output[r, c] = ExcelError.ExcelErrorNum;
                        localErrorCount++;
                        continue;
                    }

                    targets[r, c] = counts[r, c].Value;
                    candidateCount++;
                }
            }

            if (candidateCount == 0)
            {
                DerivaLog.Info(
                    "GetCurve range completed without ETTJ fetch. Cells=" +
                    (endDates.Rows * endDates.Columns) +
                    ", localErrors=" + localErrorCount);
                return output;
            }

            var preparedCurve = await EttjMemoryCache.GetPreparedCurveAsync(
                    referenceDate,
                    curve,
                    InterpolationMethod.CubicSpline)
                .ConfigureAwait(false);

            int valueCount = 0;
            int interpolationErrorCount = 0;
            for (int r = 0; r < endDates.Rows; r++)
            {
                for (int c = 0; c < endDates.Columns; c++)
                {
                    if (!targets[r, c].HasValue)
                        continue;

                    var result = MapInterpolationResult(preparedCurve.Interpolate(targets[r, c].Value));
                    output[r, c] = result;
                    if (result is double)
                        valueCount++;
                    else
                        interpolationErrorCount++;
                }
            }

            DerivaLog.Info(
                "GetCurve range completed. Curve=" + curve +
                ", referenceDate=" + referenceDate.ToString("yyyy-MM-dd") +
                ", cells=" + (endDates.Rows * endDates.Columns) +
                ", values=" + valueCount +
                ", localErrors=" + localErrorCount +
                ", interpolationErrors=" + interpolationErrorCount +
                ", curvePoints=" + preparedCurve.PointCount);

            return output;
        }

        private static object[,] ToSpillTable(IList<EttjRecord> records)
        {
            var table = new object[records.Count + 1, 7];
            table[0, 0] = "refdate";
            table[0, 1] = "curva";
            table[0, 2] = "descricao";
            table[0, 3] = "dias_corridos";
            table[0, 4] = "dias_uteis";
            table[0, 5] = "taxa";
            table[0, 6] = "vertice";

            for (int i = 0; i < records.Count; i++)
            {
                var record = records[i];
                table[i + 1, 0] = record.RefDate;
                table[i + 1, 1] = record.Curva;
                table[i + 1, 2] = record.Descricao;
                table[i + 1, 3] = record.DiasCorridos;
                table[i + 1, 4] = record.DiasUteis;
                table[i + 1, 5] = record.Taxa;
                table[i + 1, 6] = record.Vertice;
            }

            return table;
        }

        private static object MapExceptionToExcelError(Exception ex)
        {
            if (ex is EttjNoDataException)
                return ExcelError.ExcelErrorNA;

            return ExcelError.ExcelErrorValue;
        }

        private static object MapInterpolationResult(InterpolationResult result)
        {
            switch (result.Error)
            {
                case InterpolationError.None:
                    return result.Value;
                case InterpolationError.OutOfRange:
                    return ExcelError.ExcelErrorNum;
                case InterpolationError.InvalidInput:
                case InterpolationError.DuplicateX:
                default:
                    return ExcelError.ExcelErrorValue;
            }
        }

        private static int? GetTargetBusinessDays(DateTime referenceDate, DateTime endDate)
        {
            if (endDate.Date <= referenceDate.Date)
                return 0;

            return Calendar.Calendar.CountDU(referenceDate, endDate);
        }

        private static bool TryParseDate(object input, out DateTime date)
        {
            date = DateTime.MinValue;
            if (input == null ||
                input is ExcelMissing ||
                input is ExcelEmpty ||
                input is ExcelError)
                return false;

            if (input is DateTime)
            {
                date = ((DateTime)input).Date;
                return true;
            }

            if (input is double)
            {
                var serial = (double)input;
                if (serial <= 0.0)
                    return false;
                date = DateTime.FromOADate(serial).Date;
                return true;
            }

            var text = input as string;
            if (string.IsNullOrWhiteSpace(text))
                return false;

            var formats = new[] { "dd/MM/yyyy", "yyyy-MM-dd", "dd-MM-yyyy" };
            return DateTime.TryParseExact(
                text.Trim(),
                formats,
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out date);
        }

        private static bool TryParseEndDateInput(object input, out EndDateInput endDateInput)
        {
            endDateInput = null;

            var matrix = input as object[,];
            if (matrix != null)
            {
                int rows = matrix.GetLength(0);
                int cols = matrix.GetLength(1);
                var dates = new DateTime?[rows, cols];
                var valid = new bool[rows, cols];

                for (int r = 0; r < rows; r++)
                {
                    for (int c = 0; c < cols; c++)
                    {
                        DateTime parsed;
                        if (TryParseDate(matrix[r, c], out parsed))
                        {
                            dates[r, c] = parsed;
                            valid[r, c] = true;
                        }
                    }
                }

                endDateInput = EndDateInput.FromRange(dates, valid);
                return true;
            }

            DateTime scalarDate;
            if (!TryParseDate(input, out scalarDate))
                return false;

            endDateInput = EndDateInput.FromScalar(scalarDate);
            return true;
        }

        private static bool TryParseCurves(object input, out List<string> curves)
        {
            curves = null;
            try
            {
                if (input == null || input is ExcelMissing || input is ExcelEmpty)
                {
                    curves = EttjCurveCatalog.Normalize(null);
                    return true;
                }

                if (input is ExcelError)
                    return false;

                var values = new List<string>();
                var matrix = input as object[,];
                if (matrix != null)
                {
                    int rows = matrix.GetLength(0);
                    int cols = matrix.GetLength(1);
                    for (int r = 0; r < rows; r++)
                    {
                        for (int c = 0; c < cols; c++)
                        {
                            var cell = matrix[r, c];
                            if (cell == null || cell is ExcelMissing || cell is ExcelEmpty)
                                continue;
                            if (cell is ExcelError)
                                return false;
                            values.Add(cell.ToString());
                        }
                    }
                }
                else
                {
                    values.Add(input.ToString());
                }

                curves = EttjCurveCatalog.Normalize(
                    values.SelectMany(v =>
                        v.Split(new[] { ',', ';', '|', ' ', '\t', '\r', '\n' },
                            StringSplitOptions.RemoveEmptyEntries)));
                return true;
            }
            catch (EttjInvalidCurveException)
            {
                return false;
            }
        }

        private static bool TryParseSingleCurve(object input, out string curve)
        {
            curve = null;

            List<string> curves;
            if (!TryParseCurves(input, out curves))
                return false;

            if (curves == null || curves.Count != 1)
                return false;

            curve = curves[0];
            return true;
        }

        private static bool TryParseBool(object input, bool defaultValue, out bool value)
        {
            value = defaultValue;
            if (input == null || input is ExcelMissing || input is ExcelEmpty)
                return true;

            if (input is bool)
            {
                value = (bool)input;
                return true;
            }

            if (input is double)
            {
                value = Math.Abs((double)input) > double.Epsilon;
                return true;
            }

            var text = input as string;
            if (text == null)
                return false;

            switch (text.Trim().ToLowerInvariant())
            {
                case "":
                    value = defaultValue;
                    return true;
                case "true":
                case "t":
                case "yes":
                case "y":
                case "sim":
                case "s":
                case "1":
                    value = true;
                    return true;
                case "false":
                case "f":
                case "no":
                case "n":
                case "nao":
                case "0":
                    value = false;
                    return true;
                default:
                    return false;
            }
        }

        private static string BuildKey(
            DateTime startDate,
            DateTime endDate,
            IEnumerable<string> curves,
            bool useCache,
            bool ignoreErrors)
        {
            return startDate.ToString("yyyyMMdd", CultureInfo.InvariantCulture) +
                   "|" + endDate.ToString("yyyyMMdd", CultureInfo.InvariantCulture) +
                   "|" + string.Join(",", curves.ToArray()) +
                   "|" + useCache +
                   "|" + ignoreErrors;
        }

        private static string BuildGetCurveKey(
            string curve,
            DateTime referenceDate,
            EndDateInput endDateInput)
        {
            return curve +
                   "|" + referenceDate.ToString("yyyyMMdd", CultureInfo.InvariantCulture) +
                   "|" + Calendar.Calendar.Version.ToString(CultureInfo.InvariantCulture) +
                   "|" + (endDateInput.IsRange ? "R" : "S") +
                   "|" + endDateInput.Rows.ToString(CultureInfo.InvariantCulture) +
                   "x" + endDateInput.Columns.ToString(CultureInfo.InvariantCulture) +
                   "|" + HashEndDateInput(endDateInput).ToString("X16", CultureInfo.InvariantCulture);
        }

        private static ulong HashEndDateInput(EndDateInput input)
        {
            unchecked
            {
                ulong hash = 1469598103934665603UL;
                MixHash(ref hash, input.Rows);
                MixHash(ref hash, input.Columns);

                if (!input.IsRange)
                {
                    MixHash(ref hash, input.ScalarDate.Ticks);
                    return hash;
                }

                for (int r = 0; r < input.Rows; r++)
                {
                    for (int c = 0; c < input.Columns; c++)
                    {
                        MixHash(ref hash, input.IsValid(r, c) ? 1L : 0L);
                        MixHash(ref hash, input.IsValid(r, c) ? input.Dates[r, c].Value.Ticks : 0L);
                    }
                }

                return hash;
            }
        }

        private static void MixHash(ref ulong hash, long value)
        {
            unchecked
            {
                ulong data = (ulong)value;
                for (int i = 0; i < 8; i++)
                {
                    hash ^= data & 0xffUL;
                    hash *= 1099511628211UL;
                    data >>= 8;
                }
            }
        }

        private sealed class EndDateInput
        {
            private readonly bool[,] _valid;

            private EndDateInput(
                bool isRange,
                DateTime scalarDate,
                DateTime?[,] dates,
                bool[,] valid)
            {
                IsRange = isRange;
                ScalarDate = scalarDate;
                Dates = dates;
                _valid = valid;
                Rows = dates.GetLength(0);
                Columns = dates.GetLength(1);
            }

            internal bool IsRange { get; private set; }
            internal DateTime ScalarDate { get; private set; }
            internal DateTime?[,] Dates { get; private set; }
            internal int Rows { get; private set; }
            internal int Columns { get; private set; }

            internal static EndDateInput FromScalar(DateTime date)
            {
                return new EndDateInput(
                    false,
                    date,
                    new DateTime?[,] { { date } },
                    new bool[,] { { true } });
            }

            internal static EndDateInput FromRange(DateTime?[,] dates, bool[,] valid)
            {
                return new EndDateInput(true, DateTime.MinValue, dates, valid);
            }

            internal bool IsValid(int row, int column)
            {
                return _valid[row, column];
            }

            internal object[,] CreateSourceErrorMatrix(object sourceError, DateTime referenceDate)
            {
                var output = new object[Rows, Columns];
                for (int r = 0; r < Rows; r++)
                {
                    for (int c = 0; c < Columns; c++)
                    {
                        if (!IsValid(r, c))
                            output[r, c] = ExcelError.ExcelErrorValue;
                        else if (Dates[r, c].Value.Date <= referenceDate.Date)
                            output[r, c] = ExcelError.ExcelErrorNum;
                        else
                            output[r, c] = sourceError;
                    }
                }

                return output;
            }
        }
    }
}
