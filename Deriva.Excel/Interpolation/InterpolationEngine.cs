using System;

namespace Deriva.Excel.Interpolation
{
    internal enum InterpolationMethod
    {
        CubicSpline,
        Linear
    }

    internal enum InterpolationError
    {
        None,
        InvalidInput,
        DuplicateX,
        OutOfRange
    }

    internal sealed class InterpolationResult
    {
        private InterpolationResult(double value, InterpolationError error)
        {
            Value = value;
            Error = error;
        }

        internal double Value { get; private set; }
        internal InterpolationError Error { get; private set; }

        internal static InterpolationResult Success(double value)
        {
            return new InterpolationResult(value, InterpolationError.None);
        }

        internal static InterpolationResult Failure(InterpolationError error)
        {
            return new InterpolationResult(0.0, error);
        }
    }

    internal static class InterpolationEngine
    {
        internal static InterpolationResult Interpolate(
            double[] xs,
            double[] ys,
            double targetX,
            InterpolationMethod method)
        {
            PreparedInterpolator interpolator;
            var result = PreparedInterpolator.TryCreate(xs, ys, method, out interpolator);
            if (result.Error != InterpolationError.None)
                return result;

            return interpolator.Interpolate(targetX);
        }
    }
}
