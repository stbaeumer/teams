using System.Collections.Generic;

namespace teams
{
    public class Team
    {
        private string v;
        private List<Lehrer> kollegium;

        public string DisplayName { get; private set; }
        public List<Schueler> Schuelers { get; private set; }
        public List<Lehrer> Klassenleitungen { get; private set; }
        public List<Lehrer> Lehrers { get; private set; }

        public Team(string displayName, List<Schueler> schuelers, List<Lehrer> lehrers, List<Lehrer> klassenleitungen)
        {
            DisplayName = displayName;
            Schuelers = schuelers;
            Lehrers = lehrers;
            Klassenleitungen = klassenleitungen;
        }

        public Team(string displayName, List<Lehrer> kollegium)
        {
            DisplayName = displayName;
            Lehrers = kollegium;
        }
    }
}