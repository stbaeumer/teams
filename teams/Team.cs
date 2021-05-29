using System;
using System.Collections.Generic;
using System.Linq;

namespace teams
{
    public class Team
    {
        public string DisplayName { get; set; }
        public string TeamId { get; set; }
        public List<string> Owners { get; set; }
        public List<string> Members { get; set; }
        public string Kategorie { get; set; }
        /// <summary>
        /// Der Typ ist entweder 0365 oder Distribution
        /// </summary>
        public string Typ { get; set; }

        //public Team(string displayName, List<Lehrer> kollegium)
        //{
        //    DisplayName = displayName;
        //    Owners = kollegium;
        //}

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

        public Team(Teams klassenteams, string name)
        {
            Members = new List<string>();
            Owners = new List<string>();
            DisplayName = name;
            int sj = Convert.ToInt32(Global.AktSj[0].Substring(2, 2));

            if (name == "Gym13LuL")
            {
                foreach (var klassenTeam in klassenteams)
                {
                    if (klassenTeam.DisplayName.StartsWith("G") && klassenTeam.DisplayName.Contains((sj - 2).ToString()))
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
            if (name == "BlaueBriefe")
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
            if (name == "Kurs-20-Abiturienten")
            {
                foreach (var klassenTeam in klassenteams)
                {
                    if (
                            klassenTeam.DisplayName.StartsWith("G") && klassenTeam.DisplayName.Contains((sj - 2).ToString())
                        )
                    {
                        foreach (var member in klassenTeam.Members)
                        {
                            if (!this.Members.Contains(member) && member.Contains("student"))
                            {
                                this.Members.Add(member);
                            }
                        }
                    }
                }
            }
        }

        internal void TeamAnlegen(string typ)
        {
            if (Owners.Count + Members.Count > 0)
            {
                if (this.Typ == typ)
                {
                    Console.WriteLine("[+] Neue Verteilergruppe " + DisplayName);
                    Global.TeamsPs1.Add(@"    Write-Host '[+] Neue Verteilergruppe: " + DisplayName.PadRight(30) + "'");
                    Global.TeamsPs1.Add(@"    New-DistributionGroup -Name " + DisplayName + " -PrimarySmtpAddress " + DisplayName + "@berufskolleg-borken.de -Confirm:$confirm"); // -Confirm:$false
                }
                else
                {

                    if (Kategorie == "Klasse")
                    {
                        Console.WriteLine(" Die Klassengruppen müssen zuerst händisch angelegt werden. Danach bitte neu starten.");
                        Console.ReadKey();
                        Environment.Exit(0);
                    }
                }
            }
        }

        internal void OwnerUndMemberAnlegen(Team teamIst)
        {
            foreach (var sollMember in this.Members)
            {
                if (!teamIst.Members.Contains(sollMember))
                {
                    Console.WriteLine("[+] Neuer " + teamIst.Typ.Substring(0,4) + "-Member : " + sollMember.PadRight(30) + " -> " + teamIst.DisplayName);
                    Global.TeamsPs1.Add(@"    Write-Host '[+] Neuer " + teamIst.Typ.Substring(0, 4) + "-Member : " + sollMember.PadRight(30) + " -> " + teamIst.DisplayName + "'");
                    Global.TeamsPs1.Add(@"    Add-UnifiedGroupLinks -Identity " + this.TeamId + " -LinkType Member -Links '" + sollMember + "' -Confirm:$confirm");
                }
            }
            foreach (var sollOwner in this.Owners)
            {
                if (!teamIst.Owners.Contains(sollOwner))
                {
                    Console.WriteLine("[+] Neuer " + teamIst.Typ.Substring(0, 4) + "-Owner : " + sollOwner.PadRight(30) + " -> " + teamIst.DisplayName);
                    Global.TeamsPs1.Add(@"    Write-Host '[+] Neuer " + teamIst.Typ.Substring(0, 4) + "-Owner : " + sollOwner.PadRight(30) + " -> " + teamIst.DisplayName + "'");
                    Global.TeamsPs1.Add(@"    Add-UnifiedGroupLinks -Identity " + this.TeamId + " -LinkType Owner -Links '" + sollOwner + "' -Confirm:$confirm");
                }
            }
        }

        internal void OwnerUndMemberLöschen(Team teamSoll, Lehrers lehrers)
        {
            foreach (var istMember in this.Members)
            {
                if ((from l in lehrers where l.Mail == istMember select l).Any())
                {
                    if (!teamSoll.Members.Contains(istMember))
                    {
                        Console.WriteLine("[-] Owner  entfernen:" + istMember.PadRight(30) + " aus " + this.DisplayName);
                        Global.TeamsPs1.Add(@"    Write-Host '[-] Owner  entfernen: " + this.DisplayName.PadRight(30) + " aus " + this.DisplayName + "'");
                        Global.TeamsPs1.Add(@"    Remove-UnifiedGroupLinks -Identity " + this.TeamId + " -LinkType Owner -Links '" + istMember + "' -Confirm:$confirm"); // -Confirm:$false
                    }
                }
            }
            foreach (var istOwner in this.Owners)
            {
                if (!teamSoll.Owners.Contains(istOwner))
                {
                    // Nur Lehrer werden gelöscht

                    if ((from l in lehrers where l.Mail == istOwner select l).Any())
                    {
                        Console.WriteLine("[-] Owner  entfernen:" + istOwner.PadRight(30) + " aus " + this.DisplayName);
                        Global.TeamsPs1.Add(@"    Write-Host '[-] Owner  entfernen: " + this.DisplayName.PadRight(30) + " aus " + this.DisplayName + "'");
                        Global.TeamsPs1.Add(@"    Remove-UnifiedGroupLinks -Identity " + this.TeamId + " -LinkType Owner -Links '" + istOwner + "' -Confirm:$confirm"); // -Confirm:$false
                    }
                }
            }
        }

        public Team(Lehrers lehrers, string name)
        {
            this.Members = new List<string>();
            this.Owners = new List<string>();
            this.DisplayName = name.Replace("--", "-")
                                        .Replace("---", "-")
                                        .Replace("ä", "ae")
                                        .Replace("ö", "oe")
                                        .Replace("ü", "ue")
                                        .Replace("ß", "ss")
                                        .Replace("Ä", "Ae")
                                        .Replace("Ö", "Oe")
                                        .Replace("Ü", "Ue")
                                        .Replace(" ", "-")
                                        .Replace("/", "-")
                                        .Replace("---", "-");

            foreach (var lehrer in lehrers)
            {
                foreach (var anrechnung in lehrer.Anrechnungen)
                {
                    if (anrechnung.Beschr == name)
                    {
                        if (!this.Members.Contains(lehrer.Mail))
                        {
                            this.Members.Add(lehrer.Mail);
                        }                        
                    }
                }
            }
        }

        public Team(Lehrers lehrers)
        {
            this.Members = new List<string>();
            this.Owners = new List<string>();
            this.DisplayName = "Kollegium";

            foreach (var lehrer in lehrers)
            {
                this.Members.Add(lehrer.Mail);
                this.Owners.Add(lehrer.Mail);
            }
        }

        public Team()
        {
        }
    }
}