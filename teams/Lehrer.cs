using System.Collections.Generic;

namespace teams
{
    public class Lehrer
    {
        public int IdUntis { get; internal set; }
        public string Kürzel { get; internal set; }
        public string Mail { get; internal set; }
        public string Geschlecht { get; internal set; }
        public double Deputat { get; internal set; }
        public List<Anrechnung> Anrechnungen { get; internal set; }
        public string Nachname { get; internal set; }
        public string Vorname { get; internal set; }
    }
}