using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace teams
{
    public class Teams: List<Team>
    {
        public Teams(Klasses klasses, Lehrers lehrers, Schuelers schuelers, Unterrichts unterrichts)
        {
            KlassenteamsSoll(klasses, lehrers,  schuelers, unterrichts);
            KlassenleitungenSoll(klasses,lehrers);
        }

        private void KlassenleitungenSoll(Klasses klasses, Lehrers lehrers)
        {
            Team klassenleitungSoll = new Team("Klassenleitungen");
            klassenleitungSoll.Kategorie = "Klassenleitungen";

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

        private void KlassenteamsSoll(Klasses klasses, Lehrers lehrers, Schuelers schuelers, Unterrichts unterrichts)
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
                    if (item.Contains("kt153630"))
                    {
                        string a = "";
                    }
                    klassenteamSoll.Members.Add(item);
                }

                this.Add(klassenteamSoll);
            }

            Console.WriteLine("Insgesamt müssen " + this.Count + " Klassengruppen im Office365 angelegt sein.");
            File.AppendAllLines(Global.TeamsPs, new List<string>() { "# Insgesamt müssen " + this.Count + " Klassengruppen im Office365 angelegt sein." }, Encoding.UTF8);
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

                    if (zeile.Count != 4)
                    {
                        Console.WriteLine("Fehler in der Datei " + pfad + ". Falsche Spaltenzahl.");
                        Console.ReadKey();
                    }

                    // Prüfe, ob es ein Klassenteam ist

                    if ((from k in klasses where k.NameUntis == zeile[1] select k).Any())
                    {
                        AddKlasseAndMemberAndOwner(zeile, "Klasse");
                    }

                    if (zeile[1] == "Klassenleitungen")
                    {
                        AddKlasseAndMemberAndOwner(zeile, "Klassenleitungen");
                    }

                    if (zeile[1] == "Bildungsgangleitungen")
                    {
                        AddKlasseAndMemberAndOwner(zeile, "Bildungsgangleitungen");
                    }
                }
            }

            Console.WriteLine("Insgesamt " + (from d in this where d.Kategorie == "Klasse" select d).Count() + " Klassengruppen in Office365 vorhanden.");
            Console.WriteLine("Insgesamt " + (from d in this from m in d.Members where d.Kategorie == "Klassenleitungen" select m).Count() + " Klassenleitungen in der Office365-Gruppe 'Klassenleitungen' vorhanden.");
            Console.WriteLine("Insgesamt " + (from d in this from b in d.Members where d.Kategorie == "Bildungsgangleitungen" select b).Count() + " Bildungsgangleitungen in der Office365-Gruppe 'Bildungsgangleitungen' vorhanden.");

            File.AppendAllLines(Global.TeamsPs, new List<string>() { "# Insgesamt " + (from d in this where d.Kategorie == "Klasse" select d).Count() + " Klassengruppen in Office365 vorhanden." });
            File.AppendAllLines(Global.TeamsPs, new List<string>() { "# Insgesamt " + (from d in this from m in d.Members where d.Kategorie == "Klassenleitungen" select m).Count() + " Klassenleitungen in der Office365-Gruppe 'Klassenleitungen' vorhanden." });
            File.AppendAllLines(Global.TeamsPs, new List<string>() { "# Insgesamt " + (from d in this from b in d.Members where d.Kategorie == "Bildungsgangleitungen" select b).Count() + " Bildungsgangleitungen in der Office365-Gruppe 'Bildungsgangleitungen' vorhanden." });

            if (this.Count < 100)
            {
                //throw new Exception("Die Anzahl der existierenden Teams ist zu niedrig.");                
            }
        }

        private void AddKlasseAndMemberAndOwner(List<string> zeile, string kategorie)
        {
            var x = (from t in this where t.DisplayName == zeile[1] where t.TeamId == zeile[0] select t).FirstOrDefault();

            if (x == null)
            {
                if (zeile[3] == "Owner")
                {
                    this.Add(new Team(zeile[1], zeile[0], zeile[2], null, kategorie));
                }
                else
                {
                    this.Add(new Team(zeile[1], zeile[0], null, zeile[2], kategorie));
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

        internal void DoppelteKlassenFinden()
        {
            int x = 0;
            foreach (var item in this)
            {
                if (item.DisplayName == "FS20B")
                {
                    int ii = (from i in this where i.DisplayName == item.DisplayName select i).Count();
                    string a = "";
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

        internal void OwnerUndMemberLöschen(Teams klassenTeamsSoll, Lehrers lehrers)
        {
            foreach (var ts in this)
            {
                foreach (var item in ts.Owners)
                {
                    if (!item.Contains("verena.baumeister@b"))
                    {
                        // Nur Lehrer werden gelöscht

                        if ((from l in lehrers where l.Mail == item select l).Any())
                        {
                            // Jeder Owner, der im Ist aber nicht im Soll existiert, wird gelöscht

                            if (!(from o in klassenTeamsSoll where o.DisplayName == ts.DisplayName where o.Owners.Contains(item) select o.Owners.Contains(item)).Any())
                            {
                                // Referendare werden nicht gelöscht.

                                if (!(from l in lehrers where item == l.Mail where l.Kürzel.StartsWith("Y") select l).Any())
                                {
                                    Console.WriteLine("[-] Owner  entfernen:" + item.PadRight(30) + " aus " + ts.DisplayName);
                                    Global.TeamsPs1.Add(@"    Write-Host '[-] Owner  entfernen: " + item.PadRight(30) + " aus " + ts.DisplayName + "'");
                                    Global.TeamsPs1.Add(@"    Remove-UnifiedGroupLinks -Identity " + ts.TeamId + " -LinkType Owner -Links '" + item + "' -Confirm:$confirm"); // -Confirm:$false
                                }
                            }
                        }
                    }                    
                }

                foreach (var item in ts.Members)
                {
                    if (!item.Contains("verena.baumeister@b"))
                    {
                        // Member, die auch Owner sind, werden nicht gelöscht

                        if (!(from dd in klassenTeamsSoll where dd.DisplayName == ts.DisplayName where dd.Owners.Contains(item) select dd.Owners.Contains(item)).Any())
                        {
                            // Member, der im Ist aber nicht im Soll existiert, wird gelöscht 

                            if (!(from o in klassenTeamsSoll where o.DisplayName == ts.DisplayName where o.Members.Contains(item) select o).Any())
                            {
                                // Nur Lehrer werden gelöscht

                                if ((from l in lehrers where l.Mail == item select l).Any())
                                {
                                    // Referendare werden nicht gelöscht.

                                    if (!(from l in lehrers where item == l.Mail where l.Kürzel.StartsWith("Y") select l).Any())
                                    {
                                        Console.WriteLine("[-] Member entfernen:" + item.PadRight(30) + " aus " + ts.DisplayName);
                                        Global.TeamsPs1.Add(@"    Write-Host '[-] Member entfernen: " + item.PadRight(30) + " aus " + ts.DisplayName + "'");
                                        Global.TeamsPs1.Add(@"    Remove-UnifiedGroupLinks -Identity " + ts.TeamId + " -LinkType Member  -Links '" + item + "' -Confirm:$confirm");
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        internal void OwnerUndMemberAnlegen(Teams teamsIst)
        {
            foreach (var ts in this)
            {
                if (ts.Kategorie == "Klassenleitungen")
                {
                    string a = "";
                }
                var identity = (from t in teamsIst where t.DisplayName == ts.DisplayName select t.TeamId).ToList();
                
                // Für jedes Team, das den Namen der Klasse trägt, werden die Member eingefügt.

                foreach (var teamId in identity)
                {
                    // Wenn das TeamIst mehrfach mit diesem Namen vorkommt, muss für beide geprüft werden.

                    var ti = (from t in teamsIst where t.DisplayName == ts.DisplayName select t).ToList();

                    foreach (var item in ts.Members)
                    {
                        foreach (var t in ti)
                        {
                            // Jeder Member, der im Soll aber nicht im Ist existiert, wird angelegt

                            if (!t.Members.Contains(item))
                            {
                                Console.WriteLine("[+] Neuer Member : " + item.PadRight(30) + " -> " + ts.DisplayName);
                                Global.TeamsPs1.Add(@"    Write-Host '[+] Neuer Member : " + item.PadRight(30) + " -> " + ts.DisplayName + "'");
                                Global.TeamsPs1.Add(@"    Add-UnifiedGroupLinks -Identity " + teamId + " -LinkType Member -Links '" + item + "' -Confirm:$confirm");
                            }
                        }
                    }
                    foreach (var item in ts.Owners)
                    {

                        foreach (var t in ti)
                        {
                            // Jeder Owner, der im Soll aber nicht im Ist existiert, wird angelegt

                            if (!t.Owners.Contains(item))
                            {
                                Console.WriteLine("[+] Neuer Owner  : " + item.PadRight(30) + " -> " + ts.DisplayName);
                                Global.TeamsPs1.Add(@"    Write-Host '[+] Neuer Owner  : " + item.PadRight(30) + " -> " + ts.DisplayName + "'");
                                Global.TeamsPs1.Add(@"    Add-UnifiedGroupLinks -Identity " + teamId + " -LinkType Owner  -Links '" + item + "' -Confirm:$confirm");
                            }
                        }
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

                File.AppendAllLines(Global.TeamsPs, new List<string>() { "", @"Add-UnifiedGroupLinks -Identity 'Kollegium' -LinkType Members -Links '" + lehrer.Mail + "'" });
            }

            this.Add(new Team("Kollegium", kollegium));
        }
    }
}