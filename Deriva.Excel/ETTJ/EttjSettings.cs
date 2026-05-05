using System;
using System.IO;
using Newtonsoft.Json;

namespace Deriva.Excel.ETTJ
{
    internal sealed class EttjSettings
    {
        internal const string B3Source = "B3";
        internal const string B3BaseUrlTemplate =
            "https://www.b3.com.br/pesquisapregao/download?filelist=TS{YYMMDD}.ex_,";

        private const string SettingsFileName = "settings.json";

        internal static readonly string DefaultHolidayUrl =
            "https://www.anbima.com.br/feriados/arqs/feriados_nacionais.xls";

        public string HolidayUrl { get; set; }
        public string InitCurves { get; set; }
        public string EttjCacheDir { get; set; }

        internal static string AppDataDir
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "Deriva");
            }
        }

        internal static string DefaultEttjCacheDir
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "Deriva",
                    "ETTJ",
                    "Cache");
            }
        }

        internal static string SettingsFile
        {
            get { return Path.Combine(AppDataDir, SettingsFileName); }
        }

        internal static EttjSettings Load()
        {
            try
            {
                if (File.Exists(SettingsFile))
                {
                    var json = File.ReadAllText(SettingsFile);
                    var settings = JsonConvert.DeserializeObject<EttjSettings>(json);
                    if (settings != null)
                        return settings.Normalized();
                }
            }
            catch
            {
                // Fall through to defaults if settings are missing or corrupt.
            }

            return CreateDefault();
        }

        internal void Save()
        {
            var normalized = Normalized();
            Directory.CreateDirectory(AppDataDir);
            File.WriteAllText(SettingsFile, JsonConvert.SerializeObject(normalized, Formatting.Indented));
        }

        internal EttjSettings Normalized()
        {
            return new EttjSettings
            {
                HolidayUrl = string.IsNullOrWhiteSpace(HolidayUrl)
                    ? DefaultHolidayUrl
                    : HolidayUrl.Trim(),
                InitCurves = string.IsNullOrWhiteSpace(InitCurves)
                    ? "PRE,DIC"
                    : InitCurves.Trim(),
                EttjCacheDir = string.IsNullOrWhiteSpace(EttjCacheDir)
                    ? DefaultEttjCacheDir
                    : EttjCacheDir.Trim()
            };
        }

        private static EttjSettings CreateDefault()
        {
            return new EttjSettings
            {
                HolidayUrl = DefaultHolidayUrl,
                InitCurves = "PRE,DIC",
                EttjCacheDir = DefaultEttjCacheDir
            };
        }
    }
}
