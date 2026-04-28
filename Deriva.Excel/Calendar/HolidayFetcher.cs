using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using ExcelDataReader;

namespace Deriva.Excel.Calendar
{
    internal static class HolidayFetcher
    {
        private const string AnbimaUrl =
            "https://www.anbima.com.br/feriados/arqs/feriados_nacionais.xls";

        internal static void ConfigureTls()
        {
            ServicePointManager.SecurityProtocol =
                SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11;
        }

        internal static async Task<List<DateTime>> FetchAsync()
        {
            using (var client = new HttpClient())
            {
                client.Timeout = TimeSpan.FromSeconds(30);
                var bytes = await client.GetByteArrayAsync(AnbimaUrl).ConfigureAwait(false);
                return ParseXls(bytes);
            }
        }

        private static List<DateTime> ParseXls(byte[] data)
        {
            var holidays = new List<DateTime>();

            using (var stream = new MemoryStream(data))
            using (var reader = ExcelReaderFactory.CreateBinaryReader(stream))
            {
                var dataSet = reader.AsDataSet(new ExcelDataSetConfiguration
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration
                    {
                        UseHeaderRow = true
                    }
                });

                if (dataSet.Tables.Count == 0)
                    return holidays;

                var table = dataSet.Tables[0];
                foreach (System.Data.DataRow row in table.Rows)
                {
                    var cell = row[0];
                    DateTime date;

                    if (cell is DateTime dt)
                        date = dt.Date;
                    else if (cell is double serial)
                        date = DateTime.FromOADate(serial).Date;
                    else if (DateTime.TryParse(
                                 cell?.ToString(),
                                 System.Globalization.CultureInfo.InvariantCulture,
                                 System.Globalization.DateTimeStyles.None,
                                 out DateTime parsed))
                        date = parsed.Date;
                    else
                        continue;

                    holidays.Add(date);
                }
            }

            return holidays;
        }
    }
}
