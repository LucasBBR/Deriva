using System;

namespace Deriva.Excel.Interpolation
{
    internal sealed class PreparedInterpolator
    {
        private readonly double[] _xs;
        private readonly double[] _ys;
        private readonly double[] _secondDerivatives;
        private readonly InterpolationMethod _method;

        private PreparedInterpolator(
            double[] xs,
            double[] ys,
            double[] secondDerivatives,
            InterpolationMethod method)
        {
            _xs = xs;
            _ys = ys;
            _secondDerivatives = secondDerivatives;
            _method = method;
        }

        internal double MinX
        {
            get { return _xs[0]; }
        }

        internal double MaxX
        {
            get { return _xs[_xs.Length - 1]; }
        }

        internal int Count
        {
            get { return _xs.Length; }
        }

        internal static InterpolationResult TryCreate(
            double[] xs,
            double[] ys,
            InterpolationMethod method,
            out PreparedInterpolator interpolator)
        {
            interpolator = null;

            if (xs == null || ys == null)
                return InterpolationResult.Failure(InterpolationError.InvalidInput);

            int n = xs.Length;
            if (n != ys.Length || n < 2)
                return InterpolationResult.Failure(InterpolationError.InvalidInput);

            var sortedXs = (double[])xs.Clone();
            var sortedYs = (double[])ys.Clone();

            bool isSorted = true;
            for (int i = 0; i < n; i++)
            {
                if (IsInvalidNumber(sortedXs[i]) || IsInvalidNumber(sortedYs[i]))
                    return InterpolationResult.Failure(InterpolationError.InvalidInput);

                if (i > 0 && sortedXs[i] < sortedXs[i - 1])
                    isSorted = false;
            }

            if (!isSorted)
                Array.Sort(sortedXs, sortedYs);

            for (int i = 1; i < n; i++)
            {
                if (sortedXs[i] == sortedXs[i - 1])
                    return InterpolationResult.Failure(InterpolationError.DuplicateX);
            }

            var secondDerivatives = method == InterpolationMethod.CubicSpline
                ? BuildCubicSplineSecondDerivatives(sortedXs, sortedYs)
                : null;

            interpolator = new PreparedInterpolator(
                sortedXs,
                sortedYs,
                secondDerivatives,
                method);
            return InterpolationResult.Success(0.0);
        }

        internal InterpolationResult Interpolate(double targetX)
        {
            if (IsInvalidNumber(targetX))
                return InterpolationResult.Failure(InterpolationError.InvalidInput);

            if (targetX < MinX || targetX > MaxX)
                return InterpolationResult.Failure(InterpolationError.OutOfRange);

            int interval = FindInterval(_xs, targetX);
            if (targetX == _xs[interval])
                return InterpolationResult.Success(_ys[interval]);
            if (targetX == _xs[interval + 1])
                return InterpolationResult.Success(_ys[interval + 1]);

            if (_method == InterpolationMethod.Linear)
                return InterpolationResult.Success(LinearInterpolate(_xs, _ys, targetX, interval));

            return InterpolationResult.Success(CubicSplineInterpolate(
                _xs,
                _ys,
                _secondDerivatives,
                targetX,
                interval));
        }

        private static bool IsInvalidNumber(double value)
        {
            return double.IsNaN(value) || double.IsInfinity(value);
        }

        private static double[] BuildCubicSplineSecondDerivatives(double[] xs, double[] ys)
        {
            int n = xs.Length;
            var secondDerivatives = new double[n];
            var u = new double[n];

            secondDerivatives[0] = 0.0;
            u[0] = 0.0;

            for (int i = 1; i < n - 1; i++)
            {
                double sig = (xs[i] - xs[i - 1]) / (xs[i + 1] - xs[i - 1]);
                double p = sig * secondDerivatives[i - 1] + 2.0;
                secondDerivatives[i] = (sig - 1.0) / p;

                double slopeNext = (ys[i + 1] - ys[i]) / (xs[i + 1] - xs[i]);
                double slopePrevious = (ys[i] - ys[i - 1]) / (xs[i] - xs[i - 1]);
                u[i] = (6.0 * (slopeNext - slopePrevious) / (xs[i + 1] - xs[i - 1]) -
                        sig * u[i - 1]) / p;
            }

            secondDerivatives[n - 1] = 0.0;

            for (int k = n - 2; k >= 0; k--)
            {
                secondDerivatives[k] =
                    secondDerivatives[k] * secondDerivatives[k + 1] + u[k];
            }

            return secondDerivatives;
        }

        private static double LinearInterpolate(double[] xs, double[] ys, double targetX, int interval)
        {
            double h = xs[interval + 1] - xs[interval];
            double weight = (targetX - xs[interval]) / h;
            return ys[interval] + weight * (ys[interval + 1] - ys[interval]);
        }

        private static double CubicSplineInterpolate(
            double[] xs,
            double[] ys,
            double[] secondDerivatives,
            double targetX,
            int interval)
        {
            double h = xs[interval + 1] - xs[interval];
            double a = (xs[interval + 1] - targetX) / h;
            double b = (targetX - xs[interval]) / h;

            return a * ys[interval] + b * ys[interval + 1] +
                   ((a * a * a - a) * secondDerivatives[interval] +
                    (b * b * b - b) * secondDerivatives[interval + 1]) *
                   h * h / 6.0;
        }

        private static int FindInterval(double[] xs, double targetX)
        {
            int lo = 0;
            int hi = xs.Length - 1;

            while (hi - lo > 1)
            {
                int mid = lo + (hi - lo) / 2;
                if (xs[mid] > targetX)
                    hi = mid;
                else
                    lo = mid;
            }

            return lo;
        }
    }
}
