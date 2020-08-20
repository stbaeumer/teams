using System.Collections.Generic;

namespace teams
{
    public class Team
    {        
        public string DisplayName { get; private set; }
        public List<Schueler> Schuelers { get; private set; }
        public List<Lehrer> Klassenleitungen { get; private set; }
        public List<Lehrer> Lehrers { get; private set; }
        public string TeamId { get; private set; }
        public List<string> Owners { get; set; }
        public List<string> Members { get; set; }
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

        public Team(string displayName, string teamId, string ownerMail, string memberMail)
        {
            Owners = new List<string>();
            Members = new List<string>();
            DisplayName = displayName;
            TeamId = teamId;
            if (ownerMail != null)
            {
                Owners.Add(ownerMail);
            }
            if (memberMail != null)
            {
                Members.Add(memberMail);
            }            
        }
        public Team(string displayName)
        {
            DisplayName = displayName;
            Owners = new List<string>();
            Members = new List<string>();            
        }
    }
}