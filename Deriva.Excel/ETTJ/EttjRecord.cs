using System;

namespace Deriva.Excel.ETTJ
{
    internal sealed class EttjRecord
    {
        public DateTime RefDate { get; set; }
        public string Curva { get; set; }
        public string Descricao { get; set; }
        public int DiasCorridos { get; set; }
        public int DiasUteis { get; set; }
        public double Taxa { get; set; }
        public string Vertice { get; set; }
    }
}
