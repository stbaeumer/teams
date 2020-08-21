using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace teams
{
    public class Teams: List<Team>
    {
        public Teams(Klasses klasses, Lehrers lehrers, Schuelers schuelers, Unterrichts unterrichts)
        {
            foreach (var klasse in (from k in klasses where (k.Klassenleitungen != null && k.Klassenleitungen.Count > 0 && k.Klassenleitungen[0] != null) select k))
            {
                Team klassenteamSoll = new Team(klasse.NameUntis);

                var o = (from l in lehrers
                         where ((from u in unterrichts where u.KlasseKürzel == klasse.NameUntis select u.LehrerKürzel).ToList()).Contains(l.Kürzel)
                         where l.Mail != null
                         where l.Mail != ""
                         select l.Mail).ToList();

                foreach (var item in o)
                {
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

            Console.WriteLine("Insgesamt " + this.Count + " Klassengruppen müssen im Office365 vorhanden sein.");
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
                    zeile.AddRange(currentLine.Replace("\"", "").Replace("\\", "").Split(','));

                    // Prüfe, ob es ein Klassenteam ist

                    if ((from k in klasses where k.NameUntis == zeile[1] select k).Any())
                    {
                        var x = (from t in this where t.DisplayName == zeile[1] select t).FirstOrDefault();

                        if (x == null)
                        {
                            if (zeile[3] == "Owner")
                            {
                                this.Add(new Team(zeile[1], zeile[0], zeile[2], null));
                            }
                            else
                            {
                                this.Add(new Team(zeile[1], zeile[0], null, zeile[2]));
                            }
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
            }
            Console.WriteLine("Insgesamt " + this.Count + " Klassengruppen in Office365 vorhanden.");
        }

        internal void OwnerUndMemberLöschen(Teams klassenTeamsSoll)
        {
            foreach (var ts in this)
            {
                var identity = (from t in this where t.DisplayName == ts.DisplayName select t.TeamId).FirstOrDefault();

                if(identity == null || identity == "")
                {
                    string a = "";
                }

                foreach (var item in ts.Members)
                {
                    // Jeder Member, der im Ist aber nicht im Soll existiert, wird gelöscht 

                    if (!(from o in klassenTeamsSoll where o.Members.Contains(item) select o).Any())
                    {
                        Console.WriteLine("[-] Member entfernen:" + item.PadRight(30) + " aus " + ts.DisplayName);
                        Global.TeamsPs1.Add(@"Write-Host '[-] Member entfernen: " + item.PadRight(30) + " aus " + ts.DisplayName + "'");
                        Global.TeamsPs1.Add(@"Remove-UnifiedGroupLinks -Identity " + ts.DisplayName.PadRight(6) + " -LinkType Member  -Links '" + item + "'");
                    }
                }

                foreach (var item in ts.Owners)
                {
                    // Jeder Owner, der im Ist aber nicht im Soll existiert, wird gelöscht

                    if (!(from o in klassenTeamsSoll where o.Owners.Contains(item) select o).Any())
                    {
                        Console.WriteLine("[-] Owner  entfernen:" + item.PadRight(30) + " aus " + ts.DisplayName);
                        Global.TeamsPs1.Add(@"Write-Host '[-] Owner  entfernen:" + item.PadRight(30) + " aus " + ts.DisplayName + "'");
                        Global.TeamsPs1.Add(@"Remove-UnifiedGroupLinks -Identity " + ts.DisplayName.PadRight(6) + " -LinkType Owner -Links '" + item + "'");
                    }
                }                
            }
        }

        internal void OwnerUndMemberAnlegen(Teams teamsIst)
        {
            foreach (var ts in this)
            {
                var identity = (from t in teamsIst where t.DisplayName == ts.DisplayName select t.TeamId).FirstOrDefault();

                if (identity == null || identity == "")
                {
                    string a = "";
                }

                foreach (var item in ts.Owners)
                {
                    // Jeder Owner, der im Soll aber nicht im Ist existiert, wird angelegt

                    if (!(from o in teamsIst where o.Owners.Contains(item) select o).Any())
                    {
                        Console.WriteLine("[+] Neuer Owner  : " + item.PadRight(30) + " -> " + ts.DisplayName);
                        Global.TeamsPs1.Add(@"Write-Host '[+] Neuer Owner: " + item.PadRight(30) + " -> " + ts.DisplayName + "'");
                        Global.TeamsPs1.Add(@"Add-UnifiedGroupLinks -Identity " + ts.DisplayName.PadRight(6) + " -LinkType Owner  -Links '" + item + "'");
                    }
                }
                foreach (var item in ts.Members)
                {
                    // Jeder Member, der im Soll aber nicht im Ist existiert, wird angelegt

                    if (!(from o in teamsIst where o.Members.Contains(item) select o).Any())
                    {
                        Console.WriteLine("[+] Neuer Member : " + item.PadRight(30) + " -> " + ts.DisplayName);
                        Global.TeamsPs1.Add(@"Write-Host '[+] Neuer Member : " + item.PadRight(30) + " -> " + ts.DisplayName + "'");
                        Global.TeamsPs1.Add(@"Add-UnifiedGroupLinks -Identity " + ts.DisplayName.PadRight(6) + " -LinkType Member -Links '" + item + "'");
                    }
                }
            }
        }

        internal void FehlendeKlassenAnlegen(Teams teamsIst)
        {
            Console.WriteLine("Fehlende Klassenteams werden angelegt.");

            int fehlend = 0;

            foreach (var item in this)
            {
                if (!(from i in teamsIst where i.DisplayName == item.DisplayName select i).Any())
                {
                    Console.WriteLine("[+] Neue Klassengruppe " + item.DisplayName);
                    fehlend++;
                }
            }
            if (fehlend > 0)
            {
                Console.WriteLine(" Die Klassengruppen müssen zuerst händisch angelegt werden. Danach geht es weiter.");
                Console.ReadKey();
                Environment.Exit(0);
            }
        }

        public string Pfad { get; private set; }

        internal void Kollegiumsteam(Lehrers lehrers)
        {
            List<Lehrer> kollegium = new List<Lehrer>();

            foreach (var lehrer in lehrers)
            {   
                kollegium.Add(lehrer);

                // Lehrer zur Office365-Gruppe hinzufügen

                File.AppendAllText(Global.TeamsPs,@"Add-UnifiedGroupLinks -Identity 'Kollegium' -LinkType Members -Links '" + lehrer.Mail + "'");
            }

            this.Add(new Team("Kollegium", kollegium));
        }
    }
}