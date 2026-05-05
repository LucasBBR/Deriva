using System;

namespace Deriva.Excel.ETTJ
{
    internal class EttjException : Exception
    {
        internal EttjException(string message) : base(message) { }
        internal EttjException(string message, Exception innerException) : base(message, innerException) { }
    }

    internal sealed class EttjNoDataException : EttjException
    {
        internal EttjNoDataException(string message) : base(message) { }
    }

    internal sealed class EttjInvalidCurveException : EttjException
    {
        internal EttjInvalidCurveException(string curve)
            : base("Invalid ETTJ curve: " + curve)
        {
            Curve = curve;
        }

        internal string Curve { get; private set; }
    }

    internal sealed class EttjParsingException : EttjException
    {
        internal EttjParsingException(string message) : base(message) { }
        internal EttjParsingException(string message, Exception innerException) : base(message, innerException) { }
    }

    internal sealed class EttjNetworkException : EttjException
    {
        internal EttjNetworkException(string message) : base(message) { }
        internal EttjNetworkException(string message, Exception innerException) : base(message, innerException) { }
    }
}
