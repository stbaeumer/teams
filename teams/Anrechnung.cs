using System;

namespace teams
{
    public class Anrechnung
    {
        public int LehrerIdUntis { get; set; }
        public bool Gelöscht { get; set; }
        public double Wert { get; set; }
        public string Beschreibung { get; set; }
        public string Grund { get; set; }
        public DateTime Von { get; set; }
        public DateTime Bis { get; set; }
        public string Zeitart { get; internal set; }
        public string Kategorie { get; internal set; }
        public int TeacherIdUntis { get; internal set; }
    }
}