using System.Collections.Generic;

namespace teams
{
    public class Team
    {        
        public string DisplayName { get; set; }
        public List<Schueler> Schuelers { get; private set; }
        public List<Lehrer> Klassenleitungen { get; private set; }
        public List<Lehrer> Lehrers { get; private set; }
        public string TeamId { get; private set; }
        public List<string> Owners { get; set; }
        public List<string> Members { get; set; }
        public string Kategorie { get; set; }
        /// <summary>
        /// Der Typ ist entweder 0365 oder Distribution
        /// </summary>
        public string Typ { get; private set; }
        public Teams Klassenteams { get; }
        public string V1 { get; }
        public string V2 { get; }

        public Team(string displayName, List<Schueler> schuelers, List<Lehrer> lehrers, List<Lehrer> klassenleitungen, string kategorie, string x)
        {
            DisplayName = displayName;
            Schuelers = schuelers;
            Lehrers = lehrers;
            Klassenleitungen = klassenleitungen;
            Kategorie = kategorie;
        }

        public Team(string displayName, List<Lehrer> kollegium)
        {
            DisplayName = displayName;
            Lehrers = kollegium;
        }

        public Team(string displayName, string teamId, string ownerMail, string memberMail, string typ, string kategorie)
        {
            Owners = new List<string>();
            Members = new List<string>();
            DisplayName = displayName;
            TeamId = teamId;
            Kategorie = kategorie;
            Typ = typ;
            
            if (ownerMail != null)
            {
                Owners.Add(ownerMail);
            }
            if (memberMail != null)
            {
                Members.Add(memberMail);
            }            
        }

        public Team(string displayName, string kategorie)
        {
            this.DisplayName = displayName;
            Kategorie = kategorie;
            Owners = new List<string>();
            Members = new List<string>();            
        }

        public Team(Teams klassenteams, int sj, string name)
        {
            Members = new List<string>();
            DisplayName = name;

            if (name == "Gym13LuL")
            {
                foreach (var klassenTeam in klassenteams)
                {
                    if (klassenTeam.DisplayName.StartsWith("G") && klassenTeam.DisplayName.Contains((sj -2).ToString()))
                    {
                        foreach (var member in klassenTeam.Members)
                        {
                            if (!this.Members.Contains(member) && !member.Contains("student"))
                            {
                                this.Members.Add(member);
                            }
                        }
                    }
                }
            }

            if (name == "AnlageBCAbschlussklassenLuL")
            {
                foreach (var klassenTeam in klassenteams)
                {
                    if (
                            klassenTeam.DisplayName.StartsWith("BS") && klassenTeam.DisplayName.Contains((sj - 1).ToString()) ||
                            klassenTeam.DisplayName.StartsWith("HB") && klassenTeam.DisplayName.Contains((sj - 1).ToString()) ||
                            klassenTeam.DisplayName.StartsWith("FM") && klassenTeam.DisplayName.Contains((sj).ToString()) ||
                            klassenTeam.DisplayName.StartsWith("FS") && klassenTeam.DisplayName.Contains((sj - 1).ToString())
                        )
                    {
                        foreach (var member in klassenTeam.Members)
                        {
                            if (!this.Members.Contains(member) && !member.Contains("student"))
                            {
                                this.Members.Add(member);
                            }
                        }
                    }
                }
            }
            if (name=="BlaueBriefe")
            {
                foreach (var klassenTeam in klassenteams)
                {
                    if (
                            klassenTeam.DisplayName.StartsWith("G") && klassenTeam.DisplayName.Contains((sj).ToString()) ||
                            klassenTeam.DisplayName.StartsWith("HB") && klassenTeam.DisplayName.Contains((sj).ToString())                            
                        )
                    {
                        foreach (var member in klassenTeam.Members)
                        {
                            if (!this.Members.Contains(member) && !member.Contains("student"))
                            {
                                this.Members.Add(member);
                            }
                        }
                    }
                }
            }
        }

        public Team(Lehrers lehrers, string name)
        {
            this.Members = new List<string>();
            this.DisplayName = name;

            foreach (var lehrer in lehrers)
            {
                foreach (var anrechnung in lehrer.Anrechnungen)
                {
                    if (anrechnung.Kategorie == name)
                    {
                        if (!this.Members.Contains(lehrer.Mail))
                        {
                            this.Members.Add(lehrer.Mail);
                        }                        
                    }
                }
            }
        }
    }
}