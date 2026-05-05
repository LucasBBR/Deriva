using System;
using Deriva.Excel.Interpolation;
using ExcelDna.Integration;

namespace Deriva.Excel.Functions
{
    public static class InterpolationFunctions
    {
        [ExcelFunction(
            Name         = "InterpolateValues",
            Description  = "Interpola valores por cubic spline natural por padrão, ou linear quando method = \"linear\".",
            Category     = "Deriva - Interpolação",
            IsThreadSafe = true)]
        public static object InterpolateValues(
            [ExcelArgument(Name = "x_values", Description = "Range com os valores do eixo X, aceitando números ou datas.")]
            object xValues,
            [ExcelArgument(Name = "y_values", Description = "Range com os valores do eixo Y.")]
            object yValues,
            [ExcelArgument(Name = "x", Description = "Valor de X para interpolar.")]
            object targetX,
            [ExcelArgument(Name = "method", Description = "Opcional: cubic, cubic_spline, spline ou linear. Default = cubic spline.")]
            object method = null)
        {
            double[] xs;
            double[] ys;
            double x;
            InterpolationMethod interpolationMethod;

            if (!TryFlattenNumericRange(xValues, out xs) ||
                !TryFlattenNumericRange(yValues, out ys) ||
                !TryConvertToDouble(targetX, out x) ||
                !TryParseMethod(method, out interpolationMethod))
                return ExcelError.ExcelErrorValue;

            var result = InterpolationEngine.Interpolate(xs, ys, x, interpolationMethod);
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

        private static bool TryFlattenNumericRange(object input, out double[] values)
        {
            values = null;

            var matrix = input as object[,];
            if (matrix != null)
            {
                int rows = matrix.GetLength(0);
                int cols = matrix.GetLength(1);
                int total = rows * cols;

                if (total == 0)
                    return false;

                values = new double[total];
                int index = 0;

                for (int r = 0; r < rows; r++)
                {
                    for (int c = 0; c < cols; c++)
                    {
                        if (!TryConvertToDouble(matrix[r, c], out values[index]))
                            return false;
                        index++;
                    }
                }

                return true;
            }

            double value;
            if (!TryConvertToDouble(input, out value))
                return false;

            values = new[] { value };
            return true;
        }

        private static bool TryConvertToDouble(object input, out double value)
        {
            value = 0.0;

            if (input == null ||
                input is ExcelMissing ||
                input is ExcelEmpty ||
                input is ExcelError)
                return false;

            if (input is DateTime)
            {
                value = ((DateTime)input).ToOADate();
                return true;
            }

            if (input is double)
            {
                value = (double)input;
                return IsValidNumber(value);
            }

            if (input is int)
            {
                value = (int)input;
                return true;
            }

            if (input is decimal)
            {
                value = Convert.ToDouble((decimal)input);
                return IsValidNumber(value);
            }

            if (input is float)
            {
                value = (float)input;
                return IsValidNumber(value);
            }

            if (input is long)
            {
                value = (long)input;
                return true;
            }

            return false;
        }

        private static bool TryParseMethod(object input, out InterpolationMethod method)
        {
            method = InterpolationMethod.CubicSpline;

            if (input == null || input is ExcelMissing || input is ExcelEmpty)
                return true;

            var raw = input as string;
            if (raw == null)
                return false;

            switch (raw.Trim().ToLowerInvariant())
            {
                case "":
                case "cubic":
                case "cubic_spline":
                case "spline":
                    method = InterpolationMethod.CubicSpline;
                    return true;
                case "linear":
                    method = InterpolationMethod.Linear;
                    return true;
                default:
                    return false;
            }
        }

        private static bool IsValidNumber(double value)
        {
            return !double.IsNaN(value) && !double.IsInfinity(value);
        }
    }
}
