using System;
using System.Collections.Generic;
using ExcelDna.Integration;

namespace Deriva.Excel.Functions
{
    public static class CalendarFunctions
    {
        [ExcelFunction(
            Name         = "IsDU",
            Description  = "Retorna VERDADEIRO se a data for um Dia Útil (não é fim de semana nem feriado nacional ANBIMA).",
            Category     = "Deriva - Calendário",
            IsThreadSafe = true)]
        public static object IsDU(
            [ExcelArgument(Name = "data", Description = "A data a verificar.")]
            DateTime date)
        {
            var result = Calendar.Calendar.IsDU(date);
            if (result == null)
                return ExcelError.ExcelErrorNA;
            return result.Value;
        }

        [ExcelFunction(
            Name         = "IsHoliday",
            Description  = "Retorna VERDADEIRO se a data for um feriado nacional ANBIMA. Fins de semana NÃO são feriados.",
            Category     = "Deriva - Calendário",
            IsThreadSafe = true)]
        public static object IsHoliday(
            [ExcelArgument(Name = "data", Description = "A data a verificar.")]
            DateTime date)
        {
            var result = Calendar.Calendar.IsHoliday(date);
            if (result == null)
                return ExcelError.ExcelErrorNA;
            return result.Value;
        }

        [ExcelFunction(
            Name         = "DU",
            Description  = "Retorna o número de dias úteis entre duas datas (início exclusivo, fim inclusivo). Ordem não importa; resultado sempre positivo.",
            Category     = "Deriva - Calendário",
            IsThreadSafe = true)]
        public static object DU(
            [ExcelArgument(Name = "data1", Description = "Primeira data.")]
            DateTime date1,
            [ExcelArgument(Name = "data2", Description = "Segunda data.")]
            DateTime date2)
        {
            var result = Calendar.Calendar.CountDU(date1, date2);
            if (result == null)
                return ExcelError.ExcelErrorNA;
            return result.Value;
        }

        [ExcelFunction(
            Name         = "ProxDU",
            Description  = "Desloca uma data por N dias úteis. N=0 retorna a própria data se for DU, caso contrário o próximo DU.",
            Category     = "Deriva - Calendário",
            IsThreadSafe = true)]
        public static object ProxDU(
            [ExcelArgument(Name = "data_inicial", Description = "Data de partida.")]
            DateTime initialDate,
            [ExcelArgument(Name = "dias_uteis", Description = "Número de dias úteis a deslocar (positivo = futuro, negativo = passado, zero = ajuste F).")]
            int businessDays)
        {
            var result = Calendar.Calendar.ShiftDU(initialDate, businessDays);
            if (result == null)
                return ExcelError.ExcelErrorNA;
            return result.Value;
        }

        [ExcelFunction(
            Name         = "AdjustDU",
            Description  = "Ajusta uma data para o dia útil mais próximo conforme a convenção: F, MF, P, MP.",
            Category     = "Deriva - Calendário",
            IsThreadSafe = true)]
        public static object AdjustDU(
            [ExcelArgument(Name = "data", Description = "A data a ajustar.")]
            DateTime date,
            [ExcelArgument(Name = "convenção", Description = "F = Following, MF = Modified Following, P = Preceding, MP = Modified Preceding.")]
            string convention)
        {
            var conv = convention?.Trim().ToUpperInvariant();
            if (conv != "F" && conv != "MF" && conv != "P" && conv != "MP")
                return ExcelError.ExcelErrorValue;

            var result = Calendar.Calendar.AdjustToConvention(date, conv);
            if (result == null)
                return ExcelError.ExcelErrorNA;
            return result.Value;
        }

        [ExcelFunction(
            Name         = "ProxMonths",
            Description  = "Adiciona meses a uma data e ajusta ao dia útil conforme a convenção (F, MF, P, MP).",
            Category     = "Deriva - Calendário",
            IsThreadSafe = true)]
        public static object ProxMonths(
            [ExcelArgument(Name = "data", Description = "Data de partida.")]
            DateTime date,
            [ExcelArgument(Name = "meses", Description = "Número de meses a adicionar (pode ser negativo).")]
            int months,
            [ExcelArgument(Name = "convenção", Description = "F = Following, MF = Modified Following, P = Preceding, MP = Modified Preceding.")]
            string convention)
        {
            var conv = convention?.Trim().ToUpperInvariant();
            if (conv != "F" && conv != "MF" && conv != "P" && conv != "MP")
                return ExcelError.ExcelErrorValue;

            var shifted = date.AddMonths(months);
            var result = Calendar.Calendar.AdjustToConvention(shifted, conv);
            if (result == null)
                return ExcelError.ExcelErrorNA;
            return result.Value;
        }

        [ExcelFunction(
            Name         = "Holidays",
            Description  = "Retorna uma matriz vertical com todos os feriados nacionais ANBIMA entre as duas datas (inclusivo).",
            Category     = "Deriva - Calendário",
            IsThreadSafe = true)]
        public static object Holidays(
            [ExcelArgument(Name = "data_inicio", Description = "Data de início do período.")]
            DateTime startDate,
            [ExcelArgument(Name = "data_fim", Description = "Data de fim do período.")]
            DateTime endDate)
        {
            List<DateTime> result = Calendar.Calendar.GetHolidays(startDate, endDate);
            if (result == null || result.Count == 0)
                return ExcelError.ExcelErrorNA;

            var arr = new object[result.Count, 1];
            for (int i = 0; i < result.Count; i++)
                arr[i, 0] = result[i];
            return arr;
        }
    }
}
