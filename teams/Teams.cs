using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace teams
{
    public class Teams: List<Team>    
    {                
        public Teams(Klasses klasses, Lehrers lehrers)
        {
            Team klassenleitungSoll = new Team("Klassenleitungen", "Klassenleitungen");

            klassenleitungSoll.Owners.Add("stefan.baeumer@berufskolleg-borken.de");

            foreach (var klassenleitungen in (from k in klasses where (k.Klassenleitungen != null && k.Klassenleitungen.Count > 0 && k.Klassenleitungen[0] != null) select k.Klassenleitungen))
            {
                foreach (var klassenleitung in klassenleitungen)
                {
                    if (!klassenleitungSoll.Members.Contains(klassenleitung.Mail))
                    {
                        klassenleitungSoll.Members.Add(klassenleitung.Mail);
                    }
                }
            }

            this.Add(klassenleitungSoll);

            Console.WriteLine("Insgesamt müssen " + (from t in this from x in t.Members where t.Kategorie == "Klassenleitungen" select t).Count() + " Klassenleitungen im Office365 angelegt sein.");
            File.AppendAllLines(Global.TeamsPs, new List<string>() { "# Insgesamt müssen " + (from t in this from x in t.Members where t.Kategorie == "Klassenleitungen" select t).Count() + " Klassenleitungen im Office365 angelegt sein." }, Encoding.UTF8);
        }

        public Teams (Klasses klasses, Lehrers lehrers, Schuelers schuelers, Unterrichts unterrichts)
        {
            foreach (var klasse in (from k in klasses where (k.Klassenleitungen != null && k.Klassenleitungen.Count > 0 && k.Klassenleitungen[0] != null && k.NameUntis.Any(char.IsDigit)) select k))
            {
                Team klassenteamSoll = new Team(klasse.NameUntis + "-LuL", "Klasse");

                var o = (from l in lehrers
                         where ((from u in unterrichts where u.KlasseKürzel == klasse.NameUntis select u.LehrerKürzel).ToList()).Contains(l.Kürzel)
                         where l.Mail != null
                         where l.Mail != ""
                         select l.Mail).ToList();

                foreach (var item in o)
                {
                    klassenteamSoll.Members.Add(item); // Owner müssen immer auch member sein.
                    klassenteamSoll.Owners.Add(item);
                }

                var m = (from s in schuelers
                         where s.Klasse == klasse.NameUntis
                         where s.Mail != null
                         where s.Mail != ""
                         select s.Mail).ToList();

                foreach (var item in m)
                {
                    klassenteamSoll.Members.Add(item);
                }

                this.Add(klassenteamSoll);
            }

            Global.WriteLine(" Klassenteams Soll", this.Count);
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
                if (!(from istTeam in this where istTeam.DisplayName == teamSoll.DisplayName where istTeam.Typ == "Distribution" select istTeam).Any())
                {
                    teamSoll.TeamAnlegen("Distribution");
                }

                foreach (var teamIst in (from i in this where i.DisplayName == teamSoll.DisplayName where i.Typ == "Distribution" select i).ToList())
                {

                    teamSoll.OwnerUndMemberAnlegen(teamIst);
                    teamIst.OwnerUndMemberLöschen(teamSoll, lehrers);
                }
            }
            
            if (teams)
            {
                teamSoll.TeamAnlegen("O365");

                foreach (var teamIst in (from i in this where i.DisplayName == teamSoll.DisplayName where i.Typ == "O365" select i).ToList())
                {
                    teamSoll.OwnerUndMemberAnlegen(teamIst);
                    teamIst.OwnerUndMemberLöschen(teamSoll, lehrers);
                }
            }
        }


        internal void VerteilergruppenMemberLöschen(Teams teamsSoll, Lehrers lehrers)
        {
            foreach (var teamIst in (from t in this where t.Typ == "Distribution" where (t.Kategorie == "Klasse" && t.DisplayName.EndsWith("-LuL")|| t.Kategorie != "Klasse") select t).ToList())
            {
                foreach (var teamIstMember in teamIst.Members)
                {
                    // Member, der im Ist aber nicht im Soll existiert, wird gelöscht 

                    if (!(from o in teamsSoll where o.DisplayName == teamIst.DisplayName where o.Members.Contains(teamIstMember) select o).Any())
                    {
                        // Nur Lehrer werden gelöscht / Fremde Member (z.B. Schulkonferenz) werden nicht angefasst.

                        if ((from l in lehrers where l.Mail == teamIstMember select l).Any())
                        {
                            // Referendare werden nicht gelöscht.

                            if (!(from l in lehrers where teamIstMember == l.Mail where l.Kürzel.StartsWith("Y") select l).Any())
                            {
                                Console.WriteLine("[-] Member entfernen: " + teamIstMember.PadRight(30) + " aus " + teamIst.DisplayName);
                                Global.TeamsPs1.Add(@"    Write-Host '[-] Member entfernen: " + teamIstMember.PadRight(30) + " aus " + teamIst.DisplayName + "'");
                                Global.TeamsPs1.Add(@"    Remove-DistributionGroupMember -Identity " + teamIst.TeamId + " -Member '" + teamIstMember + "' -Confirm:$confirm");
                            }
                        }
                    }
                }
            }
        }

        internal void VerteilergruppeMemberAnlegen(Teams teamsIst)
        {
            foreach (var teamSoll in this)
            {
                // Bei Klassenteams wird ein LuL angehangen

                var teamIstMembers = (from t in teamsIst from m in t.Members where t.DisplayName == teamSoll.DisplayName where t.Typ == "Distribution" select m).ToList();

                foreach (var teamSollMember in teamSoll.Members)
                {
                    if (!teamSollMember.Contains("@students"))
                    {
                        // Jeder Member, der im Soll aber nicht im Ist existiert, wird angelegt

                        if (!teamIstMembers.Contains(teamSollMember.ToString()))
                        {
                            Console.WriteLine("[+] Neuer VG-Member : " + teamSollMember.PadRight(50) + " -> " + teamSoll.DisplayName);
                            Global.TeamsPs1.Add(@"    Write-Host '[+] VG-Neuer Member : " + teamSollMember.PadRight(50) + " -> " + teamSoll.DisplayName + "'");
                            Global.TeamsPs1.Add(@"    Add-DistributionGroupMember -Identity " + teamSoll.DisplayName + " -Member '" + teamSollMember + "' -Confirm:$confirm");
                        }
                    }
                }
            }
        }

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

            

            if (this.Count < 100)
            {
                //throw new Exception("Die Anzahl der existierenden Teams ist zu niedrig.");                
            }
        }

        public Teams()
        {
        }

        public Teams(Lehrers lehrers, string name)
        {
            Team teamSoll = new Team(name, name);

            teamSoll.Owners.Add("stefan.baeumer@berufskolleg-borken.de");

            if (name == "Lehrerinnen")
            {
                foreach (var lehrerin in (from l in lehrers where l.Geschlecht == "w" select l))
                {
                    teamSoll.Members.Add(lehrerin.Mail);
                }
            }
            if (name == "Vollzeit")
            {
                foreach (var vollzeitkraft in (from l in lehrers where l.Deputat == 25.5 select l))
                {
                    teamSoll.Members.Add(vollzeitkraft.Mail);
                }
            }

            if (name == "Teilzeit")
            {
                foreach (var teilzeitkraft in (from l in lehrers where l.Deputat < 25.5 select l))
                {
                    teamSoll.Members.Add(teilzeitkraft.Mail);
                }
            }

            if (name == "Bildungsgangleitung")
            {
                foreach (var lehrer in (from l in lehrers where l.Anrechnungen != null where l.Anrechnungen.Count > 0 select l))
                {
                    foreach (var anrechnung in lehrer.Anrechnungen)
                    {
                        if (anrechnung.Grund == "500" && anrechnung.Beschr == "Bildungsgangleitung")
                        {
                            teamSoll.Members.Add(lehrer.Mail);
                        }                        
                    }
                }
            }


            this.Add(teamSoll);

            Console.WriteLine("LehrerinnenSoll: " + this.Count());
            File.AppendAllLines(Global.TeamsPs, new List<string>() { "# Insgesamt " + this.Count() + " Klassenleitungen im Office365 angelegt sein." }, Encoding.UTF8);
        }

        public Teams(Teams klassenTeams)
        {
            List<string> bildungsgänge = new List<string>();

            foreach (var klassenTeam in klassenTeams)
            {
                // Die erste Ziffer und alles danach wird abgeschnitten

                var bildungsgang = klassenTeam.DisplayName.Substring(0, klassenTeam.DisplayName.IndexOfAny("0123456789".ToCharArray()));

                if (!bildungsgänge.Contains(bildungsgang))
                {
                    bildungsgänge.Add(bildungsgang);
                    
                    Team bgSoll = new Team(bildungsgang + "-LuL", "Bildungsgänge");

                    foreach (var kT in klassenTeams)
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
            Console.WriteLine("Bildungsgänge: " + this.Count());
            File.AppendAllLines(Global.TeamsPs, new List<string>() { "# Insgesamt " + this.Count() + " Klassenleitungen im Office365 angelegt sein." }, Encoding.UTF8);
        }

        public Teams(Lehrers lehrers)
        {
            Lehrers = lehrers;
        }

        internal void DoppelteKlassenFinden()
        {
            int x = 0;
            foreach (var item in this)
            {
                if (item.DisplayName == "FS20B")
                {
                    int ii = (from i in this where i.DisplayName == item.DisplayName select i).Count();                    
                }
                if ((from i in this where i.DisplayName == item.DisplayName select i).Count() > 1)
                {
                    Console.WriteLine("Die Klasse " + item.DisplayName + " ist mehrfach als O365-Gruppe angelegt.");
                    x++;
                }
            }
            if (x > 0)
            {
                Console.WriteLine("ENTER");

                Console.ReadKey();
            }
        }

        



        public Lehrers Lehrers { get; }

        //internal void Kollegiumsteam(Lehrers lehrers)
        //{
        //    List<Lehrer> kollegium = new List<Lehrer>();

        //    foreach (var lehrer in lehrers)
        //    {   
        //        kollegium.Add(lehrer);

        //        // Lehrer zur Office365-Gruppe hinzufügen

        //        File.AppendAllLines(Global.TeamsPs, new List<string>() { "", @"Add-UnifiedGroupLinks -Identity 'Kollegium' -LinkType Members -Links '" + lehrer.Mail + "'" });
        //    }

        //    this.Add(new Team("Kollegium", kollegium));
        //}
    }
}