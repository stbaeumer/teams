using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace teams
{
    public class Teams : List<Team>
    {
        /// <summary>
        /// TeamsIst
        /// </summary>
        /// <param name="pfad"></param>
        /// <param name="klasses"></param>
        public Teams(string pfad, Klasses klasses)
        {
            using (StreamReader reader = new StreamReader(pfad))
            {
                string currentLine;

                reader.ReadLine();

                while ((currentLine = reader.ReadLine()) != null)
                {
                    List<string> zeile = new List<string>();

                    zeile.AddRange(currentLine.Replace("\"", "").Replace("\\", "").Split('|'));

                    if (zeile.Count != 5)
                    {
                        Console.WriteLine("Fehler in der Datei " + pfad + ". Falsche Spaltenzahl.");
                        Console.ReadKey();
                    }

                    var x = (from t in this where t.DisplayName == zeile[1] where t.TeamId == zeile[0] select t).FirstOrDefault();

                    if (x == null)
                    {
                        Team teamIst = new Team();

                        teamIst.TeamId = zeile[0];
                        teamIst.DisplayName = zeile[1];
                        teamIst.Owners = new List<string>();
                        teamIst.Members = new List<string>();
                        teamIst.Typ = zeile[4];
                        teamIst.Kategorie = (from k in klasses where k.NameUntis + "-LuL" == teamIst.DisplayName select k).Any() ? "Klasse" : "";

                        if (zeile[3] == "Owner")
                        {
                            teamIst.Owners.Add(zeile[2]);
                        }
                        else
                        {
                            teamIst.Members.Add(zeile[2]);
                        }
                        this.Add(teamIst);
                    }
                    else
                    {
                        if (zeile[3] == "Owner")
                        {
                            x.Owners.Add(zeile[2]);
                        }
                        else
                        {
                            x.Members.Add(zeile[2]);
                        }
                    }
                }
            }

            Global.WriteLine("vorhandene Teams", this.Count);
            Global.WriteLine(" davon O365", (from t in this where t.Typ == "O365" select t).Count());
            Global.WriteLine(" davon Verteilergruppen", (from t in this where t.Typ == "Distribution" select t).Count());
        }

        /// <summary>
        /// KlassenteamsSoll
        /// </summary>
        /// <param name="klasses"></param>
        /// <param name="lehrers"></param>
        /// <param name="schuelers"></param>
        /// <param name="unterrichts"></param>
        public Teams(Klasses klasses, Lehrers lehrers, Schuelers schuelers, Unterrichts unterrichts)
        {
            foreach (var klasse in (from k in klasses where (k.Klassenleitungen != null && k.Klassenleitungen.Count > 0 && k.Klassenleitungen[0] != null && k.NameUntis.Any(char.IsDigit)) select k))
            {
                Team klassenteamSoll = new Team(klasse.NameUntis, "Klasse");

                klassenteamSoll.Unterrichts.AddRange((from u in unterrichts where u.KlasseKürzel == klasse.NameUntis select u).ToList());

                var unterrichtendeLehrer = (from l in lehrers
                                            where (from u in unterrichts where u.KlasseKürzel == klasse.NameUntis select u.LehrerKürzel).ToList().Contains(l.Kürzel)
                                            where l.Mail != null
                                            where l.Mail != ""
                                            select l.Mail).ToList();

                foreach (var unterrichtenderLehrer in unterrichtendeLehrer)
                {
                    klassenteamSoll.Members.Add(unterrichtenderLehrer); // Owner müssen immer auch member sein.
                    klassenteamSoll.Owners.Add(unterrichtenderLehrer);
                }

                var schuelersDerKlasse = (from s in schuelers
                                          where s.Klasse == klasse.NameUntis
                                          where s.Mail != null
                                          where s.Mail != ""
                                          select s.Mail).ToList();

                foreach (var schuelerDerKlasse in schuelersDerKlasse)
                {
                    klassenteamSoll.Members.Add(schuelerDerKlasse);
                }

                if (klassenteamSoll.Members.Count() + klassenteamSoll.Owners.Count() > 0)
                {
                    this.Add(klassenteamSoll);
                }                
            }

            Global.WriteLine("Klassenteams Soll", this.Count);
            Console.WriteLine("");
        }

        internal void SyncTeams(bool verteilergruppe, bool teams, Lehrers lehrers, Teams teamsSoll)
        {
            foreach (var teamSoll in teamsSoll)
            {
                this.SyncTeam(verteilergruppe, teams, lehrers, teamSoll);
            }
        }

        internal void SyncTeam(bool verteilergruppe, bool teams, Lehrers lehrers, Team teamSoll)
        {
            if (verteilergruppe)
            {
                Sync(lehrers, teamSoll, "Distribution");
            }

            if (teams)
            {                
                Sync(lehrers, teamSoll, "O365");
            }
        }

        private void Sync(Lehrers lehrers, Team teamSoll, string typ)
        {
            if (!(from teamIst in this where teamIst.DisplayName == teamSoll.DisplayName where teamIst.Typ == typ select teamIst).Any())
            {
                teamSoll.Typ = typ;
                teamSoll.TeamAnlegen(typ);

                // Ein neu angelegtes Team wird den Ist-Teams zugerechnet, damit anchließend direkt Member hinzugefügt werden können.
                Team t = new Team();
                t.DisplayName = teamSoll.DisplayName;
                t.Typ = typ;
                t.Members = new List<string>();
                t.Owners = new List<string>();

                this.Add(t);
            }

            foreach (var teamIst in (from i in this where i.DisplayName == teamSoll.DisplayName where i.Typ == typ select i).ToList())
            {
                teamSoll.OwnerUndMemberAnlegen(teamIst);
                teamIst.OwnerUndMemberLöschen(teamSoll, lehrers);
            }
        }

        /// <summary>
        /// Bildungsgänge
        /// </summary>
        /// <param name="klassenTeams"></param>
        /// <param name="name"></param>
        public Teams(Teams klassenTeams, string name)
        {
            Teams klassenTeamsTemp = new Teams();

            foreach (var klas in klassenTeams)
            {
                var t = new Team
                {
                    DisplayName = klas.DisplayName,
                    Members = klas.Members,
                    Owners = klas.Owners,
                    Kategorie = klas.Kategorie,
                    Typ = klas.Typ,
                    TeamId = klas.TeamId
                };
                klassenTeamsTemp.Add(t);
            }
            
           
            if (name.StartsWith("Klassenteams") && name.EndsWith("-LuL"))
            {   
                foreach (var klassenTeam in klassenTeamsTemp)
                {
                    klassenTeam.DisplayName += "-LuL";                 
                    klassenTeam.Members = (from m in klassenTeam.Members where !m.Contains("student") select m).ToList();
                    klassenTeam.Owners = new List<string>(); // Distribution Groups haben nur member
                }                

                this.AddRange(klassenTeamsTemp);
            }

            if (name.StartsWith("Klassenteams") && name.EndsWith("-SuS"))
            {
                foreach (var klassenTeam in klassenTeamsTemp)
                {
                    klassenTeam.DisplayName += "-SuS";
                    klassenTeam.Members = (from m in klassenTeam.Members where m.Contains("student") select m).ToList();
                    klassenTeam.Owners = (from m in klassenTeam.Owners where m.Contains("student") select m).ToList();
                }

                this.AddRange(klassenTeamsTemp);
            }

            if (name == "Bildungsgaenge-LuL")
            {
                List<string> bildungsgänge = new List<string>();

                foreach (var klassenTeam in klassenTeamsTemp)
                {
                    // Die erste Ziffer und alles danach wird abgeschnitten

                    var bildungsgang = klassenTeam.DisplayName.Substring(0, klassenTeam.DisplayName.IndexOfAny("0123456789".ToCharArray()));

                    if (!bildungsgänge.Contains(bildungsgang))
                    {
                        bildungsgänge.Add(bildungsgang);

                        Team bgSoll = new Team(bildungsgang + "-LuL", "Bildungsgänge");

                        foreach (var kT in klassenTeamsTemp)
                        {
                            if (kT.DisplayName.Substring(0, klassenTeam.DisplayName.IndexOfAny("0123456789".ToCharArray())) == bildungsgang)
                            {
                                foreach (var klassenTeamMember in (from k in kT.Members where !k.Contains("students") select k).ToList())
                                {
                                    if (!bgSoll.Members.Contains(klassenTeamMember) && klassenTeamMember != null && klassenTeamMember != "")
                                    {
                                        bgSoll.Members.Add(klassenTeamMember);
                                    }
                                }
                            }
                        }

                        if (bgSoll.Members.Count > 0)
                        {
                            this.Add(bgSoll);
                        }
                    }
                }
            }
        }

        public Teams()
        {
        }
    }
}