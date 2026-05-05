using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Deriva.Excel.Diagnostics;

namespace Deriva.Excel.ETTJ
{
    internal sealed class EttjClient
    {
        private const int DefaultTimeoutSeconds = 120;
        private const int DefaultRetries = 3;

        internal async Task<List<EttjRecord>> FetchAsync(
            DateTime refDate,
            IEnumerable<string> curves,
            int timeoutSeconds = DefaultTimeoutSeconds,
            int retries = DefaultRetries)
        {
            DerivaLog.Info("ETTJ FetchAsync start. RefDate=" + refDate.ToString("yyyy-MM-dd"));
            var raw = await DownloadRawAsync(refDate.Date, timeoutSeconds, retries)
                .ConfigureAwait(false);
            var text = ExtractTaxaSwapText(raw, refDate.Date);
            var records = ParseTaxaSwapText(text, curves, refDate.Date);
            DerivaLog.Info("ETTJ FetchAsync parsed rows=" + records.Count);
            return records;
        }

        internal static List<EttjRecord> ParseTaxaSwapText(
            string text,
            IEnumerable<string> curves,
            DateTime refDate)
        {
            var curveSet = curves == null
                ? null
                : new HashSet<string>(
                    curves.Select(c => c.Trim().ToUpperInvariant()),
                    StringComparer.OrdinalIgnoreCase);

            var records = new List<EttjRecord>();
            int skippedParseErrors = 0;
            int skippedShortLines = 0;
            if (string.IsNullOrEmpty(text))
                return records;

            using (var reader = new StringReader(text))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    if (line.Length < 67)
                    {
                        skippedShortLines++;
                        continue;
                    }

                    var curve = line.Substring(21, 5).Trim();
                    if (curveSet != null && !curveSet.Contains(curve))
                        continue;

                    int calendarDays;
                    int businessDays;
                    long rawRate;
                    if (!int.TryParse(line.Substring(41, 5).Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out calendarDays) ||
                        !int.TryParse(line.Substring(46, 5).Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out businessDays) ||
                        !long.TryParse(line.Substring(52, 14).Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out rawRate))
                    {
                        skippedParseErrors++;
                        continue;
                    }

                    var signChar = line[51];
                    if (signChar != '+' && signChar != '-')
                    {
                        skippedParseErrors++;
                        continue;
                    }

                    var sign = signChar == '+' ? 1.0 : -1.0;
                    records.Add(new EttjRecord
                    {
                        RefDate = refDate.Date,
                        Curva = curve,
                        Descricao = line.Substring(26, 15).Trim(),
                        DiasCorridos = calendarDays,
                        DiasUteis = businessDays,
                        Taxa = sign * rawRate / 1000000000.0,
                        Vertice = line[66].ToString()
                    });
                }
            }

            DerivaLog.Info(
                "ETTJ parse complete. RefDate=" + refDate.ToString("yyyy-MM-dd") +
                ", rows=" + records.Count +
                ", shortLines=" + skippedShortLines +
                ", parseErrors=" + skippedParseErrors);

            return records
                .OrderBy(r => r.Curva, StringComparer.Ordinal)
                .ThenBy(r => r.DiasCorridos)
                .ToList();
        }

        private static async Task<byte[]> DownloadRawAsync(
            DateTime refDate,
            int timeoutSeconds,
            int retries)
        {
            var url = BuildUrl(refDate);
            Exception lastError = null;

            for (int attempt = 1; attempt <= retries; attempt++)
            {
                try
                {
                    DerivaLog.Info("ETTJ download attempt " + attempt + "/" + retries + ". Url=" + url);
                    using (var handler = new HttpClientHandler
                    {
                        UseProxy = false,
                        AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate,
                        ServerCertificateCustomValidationCallback = (message, certificate, chain, errors) => true
                    })
                    using (var client = new HttpClient(handler))
                    {
                        client.Timeout = TimeSpan.FromSeconds(timeoutSeconds);
                        client.DefaultRequestHeaders.UserAgent.ParseAdd(
                            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36");
                        client.DefaultRequestHeaders.Accept.ParseAdd("*/*");
                        client.DefaultRequestHeaders.Referrer = new Uri(
                            "https://www.b3.com.br/pt_br/market-data-e-indices/servicos-de-dados/market-data/historico/boletins-diarios/pesquisa-por-pregao/pesquisa-por-pregao/");

                        using (var response = await client.GetAsync(url).ConfigureAwait(false))
                        {
                            DerivaLog.Info("ETTJ download response. Status=" + (int)response.StatusCode);
                            if (response.StatusCode == HttpStatusCode.ProxyAuthenticationRequired)
                                throw new EttjNetworkException("Proxy authentication required (HTTP 407).");

                            if ((int)response.StatusCode == 502 ||
                                (int)response.StatusCode == 503 ||
                                (int)response.StatusCode == 504)
                            {
                                lastError = new EttjNetworkException(
                                    "B3 server unavailable (HTTP " + (int)response.StatusCode + ").");
                                if (attempt < retries)
                                {
                                    await BackoffAsync(attempt).ConfigureAwait(false);
                                    continue;
                                }
                            }

                            if (!response.IsSuccessStatusCode)
                                throw new EttjNetworkException(
                                    "HTTP " + (int)response.StatusCode + " while accessing " + url);

                            var bytes = await response.Content.ReadAsByteArrayAsync().ConfigureAwait(false);
                            DerivaLog.Info("ETTJ download bytes=" + (bytes == null ? 0 : bytes.Length));
                            if (bytes == null || bytes.Length == 0)
                                throw new EttjNoDataException("B3 returned an empty TaxaSwap response.");

                            if (bytes.Length <= 22)
                                throw new EttjNoDataException("B3 returned an empty TaxaSwap ZIP.");

                            return bytes;
                        }
                    }
                }
                catch (TaskCanceledException ex)
                {
                    lastError = new EttjNetworkException("Timeout while accessing " + url, ex);
                    DerivaLog.Error("ETTJ download timeout.", ex);
                    if (attempt < retries)
                    {
                        await BackoffAsync(attempt).ConfigureAwait(false);
                        continue;
                    }
                }
                catch (HttpRequestException ex)
                {
                    DerivaLog.Error("ETTJ download network error.", ex);
                    throw new EttjNetworkException("Network error while accessing " + url, ex);
                }
            }

            if (lastError is EttjException)
                throw lastError;

            throw new EttjNetworkException("Unable to download TaxaSwap from B3.");
        }

        private static string BuildUrl(DateTime refDate)
        {
            return EttjSettings.B3BaseUrlTemplate.Replace(
                "{YYMMDD}",
                refDate.ToString("yyMMdd", CultureInfo.InvariantCulture));
        }

        private static async Task BackoffAsync(int attempt)
        {
            await Task.Delay(TimeSpan.FromSeconds(Math.Pow(2.0, attempt))).ConfigureAwait(false);
        }

        private static string ExtractTaxaSwapText(byte[] raw, DateTime refDate)
        {
            byte[] innerBytes;
            try
            {
                using (var outer = new ZipArchive(new MemoryStream(raw), ZipArchiveMode.Read))
                {
                    DerivaLog.Info("ETTJ outer ZIP entries=" + outer.Entries.Count);
                    if (outer.Entries.Count == 0)
                        throw new EttjParsingException("Outer TaxaSwap ZIP is empty.");

                    var outerEntry = outer.Entries.FirstOrDefault(e => !string.IsNullOrEmpty(e.Name));
                    if (outerEntry == null)
                        throw new EttjParsingException("Outer TaxaSwap ZIP has no file entries.");

                    DerivaLog.Info("ETTJ outer entry=" + outerEntry.FullName + ", compressed=" + outerEntry.CompressedLength + ", length=" + outerEntry.Length);
                    innerBytes = ReadEntryBytes(outerEntry);
                    DerivaLog.Info("ETTJ outer entry bytes=" + innerBytes.Length);
                }
            }
            catch (InvalidDataException ex)
            {
                throw new EttjParsingException("Downloaded TaxaSwap file is not a valid ZIP.", ex);
            }

            var offset = FindZipOffset(innerBytes);
            if (offset < 0)
                throw new EttjParsingException("Embedded ZIP marker was not found inside TaxaSwap .ex_ file.");
            DerivaLog.Info("ETTJ embedded ZIP offset=" + offset);

            var embeddedZipBytes = new byte[innerBytes.Length - offset];
            Buffer.BlockCopy(innerBytes, offset, embeddedZipBytes, 0, embeddedZipBytes.Length);
            DerivaLog.Info("ETTJ embedded ZIP bytes=" + embeddedZipBytes.Length);

            try
            {
                using (var innerStream = new MemoryStream(embeddedZipBytes))
                using (var innerZip = new ZipArchive(innerStream, ZipArchiveMode.Read))
                {
                    DerivaLog.Info("ETTJ inner ZIP entries=" + innerZip.Entries.Count);
                    if (innerZip.Entries.Count == 0)
                        throw new EttjParsingException("Embedded TaxaSwap ZIP is empty.");

                    var textEntry = innerZip.Entries.FirstOrDefault(e => !string.IsNullOrEmpty(e.Name));
                    if (textEntry == null)
                        throw new EttjParsingException("Embedded TaxaSwap ZIP has no file entries.");

                    DerivaLog.Info("ETTJ inner entry=" + textEntry.FullName + ", compressed=" + textEntry.CompressedLength + ", length=" + textEntry.Length);
                    var textBytes = ReadEntryBytes(textEntry);
                    DerivaLog.Info("ETTJ text bytes=" + textBytes.Length);
                    return Encoding.GetEncoding("iso-8859-1").GetString(textBytes);
                }
            }
            catch (InvalidDataException ex)
            {
                DerivaLog.Warn(
                    "ETTJ inner ZipArchive failed, trying local-header fallback. " +
                    ex.Message);
                try
                {
                    var textBytes = ExtractFirstZipEntryFromLocalHeader(embeddedZipBytes);
                    DerivaLog.Info("ETTJ local-header fallback text bytes=" + textBytes.Length);
                    return Encoding.GetEncoding("iso-8859-1").GetString(textBytes);
                }
                catch (Exception fallbackEx)
                {
                    throw new EttjParsingException(
                        "Embedded TaxaSwap ZIP is invalid for " +
                        refDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) + ".",
                        fallbackEx);
                }
            }
        }

        private static byte[] ExtractFirstZipEntryFromLocalHeader(byte[] zipBytes)
        {
            var localOffset = FindZipOffset(zipBytes);
            if (localOffset < 0)
                throw new EttjParsingException("Local ZIP header was not found.");

            if (zipBytes.Length < localOffset + 30)
                throw new EttjParsingException("Local ZIP header is truncated.");

            int flags = ReadUInt16(zipBytes, localOffset + 6);
            int method = ReadUInt16(zipBytes, localOffset + 8);
            long compressedSize = ReadUInt32(zipBytes, localOffset + 18);
            long uncompressedSize = ReadUInt32(zipBytes, localOffset + 22);
            int fileNameLength = ReadUInt16(zipBytes, localOffset + 26);
            int extraLength = ReadUInt16(zipBytes, localOffset + 28);
            int dataStart = localOffset + 30 + fileNameLength + extraLength;

            if (dataStart > zipBytes.Length)
                throw new EttjParsingException("Local ZIP file data is outside the buffer.");

            if (compressedSize == 0 || (flags & 0x08) != 0)
            {
                var centralOffset = FindSignature(zipBytes, 0x50, 0x4B, 0x01, 0x02, dataStart);
                if (centralOffset >= 0 && zipBytes.Length >= centralOffset + 46)
                {
                    compressedSize = ReadUInt32(zipBytes, centralOffset + 20);
                    uncompressedSize = ReadUInt32(zipBytes, centralOffset + 24);
                    DerivaLog.Info(
                        "ETTJ local-header fallback using central directory sizes. compressed=" +
                        compressedSize + ", uncompressed=" + uncompressedSize);
                }
                else
                {
                    throw new EttjParsingException(
                        "Local ZIP entry uses a data descriptor and no central directory sizes were found.");
                }
            }

            if (compressedSize <= 0 || dataStart + compressedSize > zipBytes.Length)
                throw new EttjParsingException("Local ZIP compressed data length is invalid.");

            DerivaLog.Info(
                "ETTJ local-header fallback. method=" + method +
                ", flags=" + flags +
                ", compressed=" + compressedSize +
                ", uncompressed=" + uncompressedSize +
                ", dataStart=" + dataStart);

            if (method == 0)
            {
                var stored = new byte[(int)compressedSize];
                Buffer.BlockCopy(zipBytes, dataStart, stored, 0, (int)compressedSize);
                return stored;
            }

            if (method != 8)
                throw new EttjParsingException("Unsupported embedded ZIP compression method: " + method);

            using (var compressed = new MemoryStream(zipBytes, dataStart, (int)compressedSize))
            using (var deflate = new DeflateStream(compressed, CompressionMode.Decompress))
            using (var output = new MemoryStream())
            {
                deflate.CopyTo(output);
                return output.ToArray();
            }
        }

        private static byte[] ReadEntryBytes(ZipArchiveEntry entry)
        {
            using (var input = entry.Open())
            using (var output = new MemoryStream())
            {
                input.CopyTo(output);
                return output.ToArray();
            }
        }

        private static int FindZipOffset(byte[] bytes)
        {
            if (bytes == null || bytes.Length < 4)
                return -1;

            for (int i = 0; i <= bytes.Length - 4; i++)
            {
                if (bytes[i] == 0x50 &&
                    bytes[i + 1] == 0x4B &&
                    bytes[i + 2] == 0x03 &&
                    bytes[i + 3] == 0x04)
                    return i;
            }

            return -1;
        }

        private static int FindSignature(
            byte[] bytes,
            byte b0,
            byte b1,
            byte b2,
            byte b3,
            int start)
        {
            if (bytes == null || bytes.Length < 4)
                return -1;

            for (int i = Math.Max(0, start); i <= bytes.Length - 4; i++)
            {
                if (bytes[i] == b0 &&
                    bytes[i + 1] == b1 &&
                    bytes[i + 2] == b2 &&
                    bytes[i + 3] == b3)
                    return i;
            }

            return -1;
        }

        private static ushort ReadUInt16(byte[] bytes, int offset)
        {
            return (ushort)(bytes[offset] | (bytes[offset + 1] << 8));
        }

        private static uint ReadUInt32(byte[] bytes, int offset)
        {
            return (uint)(bytes[offset] |
                          (bytes[offset + 1] << 8) |
                          (bytes[offset + 2] << 16) |
                          (bytes[offset + 3] << 24));
        }
    }
}
