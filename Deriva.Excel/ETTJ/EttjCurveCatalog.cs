using System;
using System.Collections.Generic;
using System.Linq;

namespace Deriva.Excel.ETTJ
{
    internal static class EttjCurveCatalog
    {
        internal const string DefaultCurve = "PRE";

        private static readonly Dictionary<string, string> Curves =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "ACC", "DIxDOL Aj. Cupom" },
                { "APR", "DIxPRE Aj. PRE" },
                { "ARB", "REAIS X ARS" },
                { "ARS", "ARS X DOL" },
                { "AUD", "DOL X AUD" },
                { "BIT", "Futuro de Bitcoin" },
                { "BRP", "IBRX50 X PRE" },
                { "CAD", "CAD X DOL" },
                { "CHF", "CHF X DOL" },
                { "CLP", "CLP X DOL" },
                { "CNH", "CNH X DOL" },
                { "CNL", "Curva Futuro CNH" },
                { "CNY", "CNY X DOL" },
                { "CYI", "CONV. YIELD" },
                { "CYM", "Convenience Yield M" },
                { "CYS", "Convenience Yield S" },
                { "CYX", "Convenience Yield X" },
                { "DCL", "CUPOM LIMPO - Dolar" },
                { "DCO", "SELIC X DOL" },
                { "DCP", "CUPOM LIMPO" },
                { "DEU", "DOL X EUR" },
                { "DGL", "Cupom Limpo de Ouro" },
                { "DIC", "DI X IPCA" },
                { "DIM", "DIxIGPM" },
                { "DOC", "DIxXDOL Cupom limpo" },
                { "DOL", "DIxDOL" },
                { "DP", "DOLxPRE" },
                { "DPL", "Cupom Limpo de Petro" },
                { "DYE", "DOL X YEN" },
                { "EBR", "Futuro de Ethereum" },
                { "ECC", "CUPOM SUJO DE EURO" },
                { "EST", "Curva Futuro Taxa" },
                { "ETR", "Futuro de Ethereum" },
                { "EUC", "CUPOM EURO" },
                { "EUR", "R$ x EURO" },
                { "FTS", "FTSE/JSE TOP 40" },
                { "GBP", "DOL X GBP" },
                { "GLD", "Curva Futuro Ouro" },
                { "HAN", "HANG SENG INDEX" },
                { "IAS", "IPCA SINTETICO" },
                { "INP", "IBOVESPA" },
                { "IPS", "IGPxPRE SINTET." },
                { "ITC", "ITC X SELIC" },
                { "JPY", "REAIS X IENE" },
                { "LEU", "Juros em EUR" },
                { "LIB", "Juros em USD" },
                { "LJP", "Juros em JPY" },
                { "MBR", "Curva de Indice BR" },
                { "MXN", "MXN X DOL" },
                { "NOK", "NOK X DOL" },
                { "NZD", "NZD X DOL" },
                { "PRE", "DIxPRE" },
                { "PTX", "PTAX" },
                { "RDA", "REAIS X DOLAR A" },
                { "RDC", "REAIS X DOLAR C" },
                { "RFS", "REAIS X FRANCO" },
                { "RLI", "REAIS X LIBRA" },
                { "RPL", "REAIS X PESO CHL" },
                { "RPM", "REAIS X PESO MEX" },
                { "RRA", "REAIS X RANDE SUL" },
                { "RUB", "RUB X DOL" },
                { "RYN", "REAIS X IUAN" },
                { "RYR", "REAIS X LIRA TUR" },
                { "RZE", "REAIS X DOLAR NZL" },
                { "SAB", "BOI GORDO" },
                { "SAC", "C.ARABICA (US$)" },
                { "SAM", "MILHO" },
                { "SAU", "SPREAD DOL AUST" },
                { "SBP", "S.BASKET X PRE" },
                { "SBR", "Futuro de Solana" },
                { "SCA", "SPREAD DOL CANADENSE" },
                { "SCF", "SPREAD FRANCO SUICO" },
                { "SCL", "SPREAD PESO CHILENO" },
                { "SCN", "SPREAD IUAN X DOL" },
                { "SDE", "DOLx EURO" },
                { "SEK", "SEK X DOL" },
                { "SFR", "Curva Futuro Taxa FR" },
                { "SGP", "SPREAD LIBRA X DOL" },
                { "SJC", "Curva de Soja CBOT" },
                { "SLP", "SELICxPRE" },
                { "SLT", "SPREAD LTN" },
                { "SML", "Curva Small Cap" },
                { "SMX", "SPREAD PESO MEX" },
                { "SNZ", "SPREAD DOLAR NZL" },
                { "SOL", "Futuro de Solana" },
                { "SOY", "Curva de Soja Futura" },
                { "STR", "SPREAD LIRA TURCA" },
                { "SYD", "SPREAD IEN X DOL" },
                { "SZA", "SPREAD RANDE SUL AFR" },
                { "TFP", "TBFxPRE" },
                { "TIC", "NTN-B" },
                { "TIE", "Curva Futuro Taxa IE" },
                { "TIM", "NTN-C" },
                { "TJP", "TJLPxPRE" },
                { "TLF", "LFT" },
                { "TP", "TRxPRE" },
                { "TPR", "LTN" },
                { "TR", "DIxTR" },
                { "TRY", "TRY X DOL" },
                { "VIX", "Curva Futuro VIX" },
                { "XFI", "IFIX" },
                { "YCC", "Cupom sujo de yen" },
                { "YCL", "IENE - Cupom Limpo" },
                { "YCS", "IEN X CUPOM" },
                { "ZAR", "ZAR X DOL" },
                { "ZEU", "Curva Zero Juro EUR" },
                { "ZMX", "Curva Zero Juro MXN" },
                { "ZUS", "Curva Zero Juro USD" },
            };

        internal static IReadOnlyDictionary<string, string> AvailableCurves
        {
            get { return Curves; }
        }

        internal static List<string> Normalize(IEnumerable<string> rawCurves)
        {
            var result = new List<string>();
            if (rawCurves != null)
            {
                foreach (var raw in rawCurves)
                {
                    if (string.IsNullOrWhiteSpace(raw))
                        continue;

                    var curve = raw.Trim().ToUpperInvariant();
                    if (curve == "TODOS" || curve == "ALL")
                        return Curves.Keys.OrderBy(k => k, StringComparer.Ordinal).ToList();

                    if (!Curves.ContainsKey(curve))
                        throw new EttjInvalidCurveException(curve);

                    if (!result.Contains(curve))
                        result.Add(curve);
                }
            }

            if (result.Count == 0)
                result.Add(DefaultCurve);

            return result;
        }
    }
}
