using Deriva.Excel.Interpolation;

namespace Deriva.Excel.ETTJ
{
    internal sealed class PreparedEttjCurve
    {
        private readonly PreparedInterpolator _interpolator;

        internal PreparedEttjCurve(
            string curve,
            PreparedInterpolator interpolator)
        {
            Curve = curve;
            _interpolator = interpolator;
        }

        internal string Curve { get; private set; }

        internal int PointCount
        {
            get { return _interpolator.Count; }
        }

        internal double MinBusinessDays
        {
            get { return _interpolator.MinX; }
        }

        internal double MaxBusinessDays
        {
            get { return _interpolator.MaxX; }
        }

        internal InterpolationResult Interpolate(int businessDays)
        {
            return _interpolator.Interpolate(businessDays);
        }
    }
}
