using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace Deriva.Excel.Diagnostics
{
    internal static class DerivaLog
    {
        internal static void Info(string message)
        {
            Write("INFO", message);
        }

        internal static void Warn(string message)
        {
            Write("WARN", message);
        }

        internal static void Error(string message, Exception ex = null)
        {
            Write("ERROR", ex == null ? message : message + " :: " + ex);
        }

        private static void Write(string level, string message)
        {
            var line = "[Deriva] " +
                       DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") +
                       " [" + level + "] " +
                       message;
            Debug.WriteLine(line);
            Trace.WriteLine(line);
            OutputDebugString(line);
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
        private static extern void OutputDebugString(string lpOutputString);
    }
}
