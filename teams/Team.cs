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
        public Unterrichts Unterrichts { get; private set; }
        public string Kategorie { get; set; }        
        /// <summary>
        /// Der Typ ist entweder 0365 oder Distribution
        /// </summary>
        public string Typ { get; set; }

        public Team()
        {
        }

        public Team(string displayName, string kategorie)
        {
            this.DisplayName = displayName;
            Kategorie = kategorie;
            Owners = new List<string>() { };
            Members = new List<string>() { };
            Unterrichts = new Unterrichts();

            if (Kategorie != "Klasse")
            {
                Owners.Add("stefan.baeumer@berufskolleg-borken.de");
                Members.Add("stefan.baeumer@berufskolleg-borken.de");
            }
        }

        public Team(Teams klassenteams, Klasses klasses, Lehrers lehrers, string name)
        {
            Members = new List<string>();
            Owners = new List<string>();
            DisplayName = name;

            int sj = Convert.ToInt32(Global.AktSj[0].Substring(2, 2));

            if (name == "Gym13-Klassenleitungen")
            {
                foreach (var klasse in klasses)
                {
                    if (klasse.NameUntis.StartsWith("G") && klasse.NameUntis.Contains((sj - 2).ToString()))
                    {
                        foreach (var klassenleitung in klasse.Klassenleitungen)
                        {
                            if (!this.Members.Contains(klassenleitung.Mail))
                            {
                                this.Members.Add(klassenleitung.Mail);
                            }
                        }
                    }
                }
            }

            if (name == "Gym13-SuS")
            {
                foreach (var klassenTeam in klassenteams)
                {
                    if (klassenTeam.DisplayName.StartsWith("G") && klassenTeam.DisplayName.Contains((sj - 2).ToString()))
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

            if (name == "Gym13-LuL")
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

            if (name == "AbschlussklassenBC-LuL")
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

            if (name == "AbschlussklassenBC-SuS")
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
                            if (!this.Members.Contains(member) && member.Contains("student"))
                            {
                                this.Members.Add(member);
                            }
                        }
                    }
                }
            }

            if (name == "AbschlussklassenBC-Klassenleitungen")
            {
                foreach (var klasse in klasses)
                {
                    if (
                            klasse.NameUntis.StartsWith("BS") && klasse.NameUntis.Contains((sj - 1).ToString()) ||
                            klasse.NameUntis.StartsWith("HB") && klasse.NameUntis.Contains((sj - 1).ToString()) ||
                            klasse.NameUntis.StartsWith("FM") && klasse.NameUntis.Contains((sj).ToString()) ||
                            klasse.NameUntis.StartsWith("FS") && klasse.NameUntis.Contains((sj - 1).ToString())
                        )
                    {
                        foreach (var klassenleitung in klasse.Klassenleitungen)
                        {
                            if (!this.Members.Contains(klassenleitung.Mail))
                            {
                                this.Members.Add(klassenleitung.Mail);
                            }
                        }
                    }
                }
            }

            if (name == "BlaueBriefe-LuL")
            {
                foreach (var klassenTeam in klassenteams)
                {
                    if (
                            klassenTeam.DisplayName.StartsWith("G") && klassenTeam.DisplayName.Contains((sj).ToString()) ||
                            klassenTeam.DisplayName.StartsWith("BS") && klassenTeam.DisplayName.Contains((sj).ToString()) ||
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

            // Halbjahreszeugniskonferenzen

            if (name == "Halbjahreszeugniskonferenzen BS HBG FS") // namensgleich mit Outlook-Termin
            {
                foreach (var klassenTeam in klassenteams)
                {
                    if (
                            klassenTeam.DisplayName.StartsWith("BS") ||
                            klassenTeam.DisplayName.StartsWith("FS") ||
                            klassenTeam.DisplayName.StartsWith("HBG")
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

            if (name == "Halbjahreszeugniskonferenzen BW HBW") // namensgleich mit Outlook-Termin
            {
                foreach (var klassenTeam in klassenteams)
                {
                    if (
                            klassenTeam.DisplayName.StartsWith("BW") ||
                            klassenTeam.DisplayName.StartsWith("HBW")
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

            if (name == "Halbjahreszeugniskonferenzen BT HBT FM") // namensgleich mit Outlook-Termin
            {
                foreach (var klassenTeam in klassenteams)
                {
                    if (
                            klassenTeam.DisplayName.StartsWith("BT") ||
                            klassenTeam.DisplayName.StartsWith("HBT") ||
                            klassenTeam.DisplayName.StartsWith("FM")
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

            if (name == "Halbjahreszeugniskonferenzen GE GT GW") // namensgleich mit Outlook-Termin
            {
                foreach (var klassenTeam in klassenteams)
                {
                    if (
                            klassenTeam.DisplayName.StartsWith("G")
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

            // Jahreszeugniskonferenzen

            if (name == "Jahreszeugniskonferenzen BS HBG FS") // namensgleich mit Outlook-Termin
            {
                foreach (var klassenTeam in klassenteams)
                {
                    if (
                            (klassenTeam.DisplayName.StartsWith("BS") && klassenTeam.DisplayName.Contains((sj).ToString())) ||
                            (klassenTeam.DisplayName.StartsWith("FS") && klassenTeam.DisplayName.Contains((sj).ToString())) ||                            
                            (klassenTeam.DisplayName.StartsWith("HBG") && klassenTeam.DisplayName.Contains((sj).ToString()))
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

            if (name == "Jahreszeugniskonferenzen BW HBW") // namensgleich mit Outlook-Termin
            {
                foreach (var klassenTeam in klassenteams)
                {
                    if (
                            klassenTeam.DisplayName.StartsWith("BW") ||                            
                            (klassenTeam.DisplayName.StartsWith("HBW") && klassenTeam.DisplayName.Contains((sj).ToString()))
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

            if (name == "Jahreszeugniskonferenzen BT HBT") // namensgleich mit Outlook-Termin
            {
                foreach (var klassenTeam in klassenteams)
                {
                    if (
                            klassenTeam.DisplayName.StartsWith("BT") ||
                            (klassenTeam.DisplayName.StartsWith("HBT") && klassenTeam.DisplayName.Contains((sj).ToString()))
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

            if (name == "Jahreszeugniskonferenzen GE GT GW") // namensgleich mit Outlook-Termin
            {
                foreach (var klassenTeam in klassenteams)
                {
                    if (
                            (klassenTeam.DisplayName.StartsWith("G") && !klassenTeam.DisplayName.Contains((sj - 2).ToString()))                            
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

            if (name == "Kollegium")
            {
                foreach (var lehrer in lehrers)
                {
                    if (!this.Members.Contains(lehrer.Mail))
                    {
                        this.Members.Add(lehrer.Mail);
                    }
                }
            }

            if (name == "Lehrerinnen")
            {
                foreach (var lehrer in lehrers)
                {
                    if (lehrer.Geschlecht == "w")
                    {
                        if (!this.Members.Contains(lehrer.Mail))
                        {
                            this.Members.Add(lehrer.Mail);
                        }
                    }
                }
            }

            if (name == "Vollzeit")
            {
                foreach (var vollzeitkraft in (from l in lehrers where l.Deputat == 25.5 select l))
                {
                    if (!this.Members.Contains(vollzeitkraft.Mail))
                    {
                        this.Members.Add(vollzeitkraft.Mail);
                    }
                }
            }

            if (name == "Teilzeit")
            {
                foreach (var teilzeitkraft in (from l in lehrers where l.Deputat < 25.5 select l))
                {
                    if (!this.Members.Contains(teilzeitkraft.Mail))
                    {
                        this.Members.Add(teilzeitkraft.Mail);
                    }
                }
            }
            if (name == "Klassenleitungen")
            {
                foreach (var klasse in klasses)
                {   
                    foreach (var klassenleitung in klasse.Klassenleitungen)
                    {
                        if (!this.Members.Contains(klassenleitung.Mail))
                        {
                            this.Members.Add(klassenleitung.Mail);
                        }
                    }                    
                }
            }
            if (name == "Fachschaft-Sport")
            {
                foreach (var klassenTeam in klassenteams)
                {
                    foreach (var unterricht in klassenTeam.Unterrichts)
                    {
                        if (new List<string>() { "SP", "SP G1", "SP G2" }.Contains(unterricht.FachKürzel))
                        {
                            var lehrerMail = (from l in lehrers where l.Kürzel == unterricht.LehrerKürzel select l.Mail).FirstOrDefault();

                            if (lehrerMail != null)
                            {
                                if (!this.Members.Contains(lehrerMail))
                                {
                                    this.Members.Add(lehrerMail);
                                }
                            }                            
                        }
                    }
                }
            }
            if (name == "Fachschaft-Mathe")
            {
                foreach (var klassenTeam in klassenteams)
                {
                    foreach (var unterricht in klassenTeam.Unterrichts)
                    {
                        if (new List<string>() { "M", "M FU", "M1", "M2", "M G1", "M G2", "M L1", "M L2", "M L", "ML", "ML1", "ML2" }.Contains(unterricht.FachKürzel))
                        {
                            var lehrerMail = (from l in lehrers where l.Kürzel == unterricht.LehrerKürzel select l.Mail).FirstOrDefault();

                            if (lehrerMail != null)
                            {
                                if (!this.Members.Contains(lehrerMail))
                                {
                                    this.Members.Add(lehrerMail);
                                }
                            }
                        }
                    }
                }
            }
            if (name == "Fachschaft-Deutsch")
            {
                foreach (var klassenTeam in klassenteams)
                {
                    foreach (var unterricht in klassenTeam.Unterrichts)
                    {
                        if (new List<string>() { "D", "D FU", "D1", "D2", "D G1", "D G2", "D L1", "D L2", "D L", "DL", "DL1", "DL2" }.Contains(unterricht.FachKürzel))
                        {
                            var lehrerMail = (from l in lehrers where l.Kürzel == unterricht.LehrerKürzel select l.Mail).FirstOrDefault();

                            if (lehrerMail != null)
                            {
                                if (!this.Members.Contains(lehrerMail))
                                {
                                    this.Members.Add(lehrerMail);
                                }
                            }
                        }
                    }
                }
            }
            if (name == "Fachschaft-Englisch")
            {
                foreach (var klassenTeam in klassenteams)
                {
                    foreach (var unterricht in klassenTeam.Unterrichts)
                    {
                        if (new List<string>() { "E", "E FU", "E1", "E2", "E G1", "E G2", "E L1", "E L2", "E L", "EL", "EL1", "EL2" }.Contains(unterricht.FachKürzel))
                        {
                            var lehrerMail = (from l in lehrers where l.Kürzel == unterricht.LehrerKürzel select l.Mail).FirstOrDefault();

                            if (lehrerMail != null)
                            {
                                if (!this.Members.Contains(lehrerMail))
                                {
                                    this.Members.Add(lehrerMail);
                                }
                            }
                        }
                    }
                }
            }
            if (name == "Fachschaft-Religion")
            {
                foreach (var klassenTeam in klassenteams)
                {
                    foreach (var unterricht in klassenTeam.Unterrichts)
                    {
                        if (new List<string>() { "KR", "KR FU", "KR1", "KR2", "KR G1", "KR G2", "ER", "ER G1" }.Contains(unterricht.FachKürzel))
                        {
                            var lehrerMail = (from l in lehrers where l.Kürzel == unterricht.LehrerKürzel select l.Mail).FirstOrDefault();

                            if (lehrerMail != null)
                            {
                                if (!this.Members.Contains(lehrerMail))
                                {
                                    this.Members.Add(lehrerMail);
                                }
                            }
                        }
                    }
                }
            }

            if (name == "Fachschaft-Politik-GG")
            {
                foreach (var klassenTeam in klassenteams)
                {
                    foreach (var unterricht in klassenTeam.Unterrichts)
                    {
                        if (new List<string>() { "PK", "PK FU", "PK1", "PK2", "GG G1", "GG G2" }.Contains(unterricht.FachKürzel))
                        {
                            var lehrerMail = (from l in lehrers where l.Kürzel == unterricht.LehrerKürzel select l.Mail).FirstOrDefault();

                            if (lehrerMail != null)
                            {
                                if (!this.Members.Contains(lehrerMail))
                                {
                                    this.Members.Add(lehrerMail);
                                }
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
                    Console.WriteLine("[+] Neue Verteilergruppe : " + DisplayName);
                    Global.TeamsPs1.Add(@"    Write-Host '[+] Neue Verteilergruppe: " + DisplayName.PadRight(30) + "'");
                    Global.TeamsPs1.Add(@"    New-DistributionGroup -Name " + DisplayName + " -PrimarySmtpAddress " + ToSmtp(DisplayName) + "@berufskolleg-borken.de -Confirm:$confirm"); // -Confirm:$false
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

        private string ToSmtp(string displayName)
        {
            return displayName.Replace("--", "-")
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
        }

        internal void OwnerUndMemberAnlegen(Team teamIst)
        {
            foreach (var sollMember in this.Members)
            {
                if (!teamIst.Members.Contains(sollMember))
                {
                    Console.WriteLine("    + Neuer " + teamIst.Typ.Substring(0,4) + "-Member  : " + sollMember.PadRight(44) + " -> " + teamIst.DisplayName);
                    Global.TeamsPs1.Add(@"    Write-Host '[+] Neuer " + teamIst.Typ.Substring(0, 4) + "-Member : " + sollMember.PadRight(44) + " -> " + teamIst.DisplayName + "'");

                    if (teamIst.Typ == "O365")
                    {
                        Global.TeamsPs1.Add(@"    Add-UnifiedGroupLinks -Identity " + teamIst.TeamId + " -LinkType Member -Links '" + sollMember + "' -Confirm:$confirm");
                    }
                    if (teamIst.Typ == "Distribution")
                    {
                        Global.TeamsPs1.Add(@"    Add-DistributionGroupMember -Identity " + teamIst.DisplayName + " -Member '" + sollMember + "' -Confirm:$confirm");
                    }
                }
            }
            foreach (var sollOwner in this.Owners)
            {
                if (!teamIst.Owners.Contains(sollOwner))
                {
                    Console.WriteLine("    + Neuer " + teamIst.Typ.Substring(0, 4) + "-Owner  :  " + sollOwner.PadRight(44) + " -> " + teamIst.DisplayName);
                    Global.TeamsPs1.Add(@"    Write-Host '[+] Neuer " + teamIst.Typ.Substring(0, 4) + "-Owner : " + sollOwner.PadRight(44) + " -> " + teamIst.DisplayName + "'");
                    if (teamIst.Typ == "O365")
                    {
                        Global.TeamsPs1.Add(@"    Add-UnifiedGroupLinks -Identity " + teamIst.TeamId + " -LinkType Owner -Links '" + sollOwner + "' -Confirm:$confirm");
                    }

                    // In Verteilergruppen werden alle Owner zu Member

                    if (teamIst.Typ == "Distribution" && !teamIst.Members.Contains(sollOwner))
                    {
                        Global.TeamsPs1.Add(@"    Add-DistributionGroupMember -Identity " + teamIst.DisplayName + " -Member '" + sollOwner + "' -Confirm:$confirm");
                    }
                }
            }
        }

        internal void OwnerUndMemberLöschen(Team teamSoll, Lehrers lehrers)
        {
            foreach (var istMember in this.Members)
            { 
                if (!teamSoll.Members.Contains(istMember))
                {
                    // Schüler werden bedingungslos gelöscht.
                    // Lehrer werden nur im August, September und nach den HZ-Zeugniskonferenzen gelöscht.
                    // Andere Member (Sekretariat usw.) werden nie gelöscht.

                    if (istMember.Contains("students") || (lehrers.istLehrer(istMember) && (DateTime.Now.Month == 9 || DateTime.Now.Month == 8 || (DateTime.Now.Month == 2 && DateTime.Now.Day > 8))))
                    {
                        Console.WriteLine("    - Owner entfernen :" + istMember.PadRight(30) + " aus " + this.DisplayName);
                        Global.TeamsPs1.Add(@"    Write-Host '[-] Owner  entfernen: " + this.DisplayName.PadRight(30) + " aus " + this.DisplayName + "'");
                        Global.TeamsPs1.Add(@"    Remove-UnifiedGroupLinks -Identity " + this.TeamId + " -LinkType Owner -Links '" + istMember + "' -Confirm:$confirm"); // -Confirm:$false
                    }
                }
            }
            foreach (var istOwner in this.Owners)
            {
                if (!teamSoll.Owners.Contains(istOwner))
                {
                    if ((DateTime.Now.Month == 9 || DateTime.Now.Month == 8 || (DateTime.Now.Month == 2 && DateTime.Now.Day > 8)) && !istOwner.Contains("students"))
                    {
                        // Lehrer werden nur im August, September und nach den HZ-Zeugniskonferenzen gelöscht.

                        if ((from l in lehrers where l.Mail == istOwner select l).Any())
                        {
                            Console.WriteLine("      - Owner  entfernen:" + istOwner.PadRight(30) + " aus " + this.DisplayName);
                            Global.TeamsPs1.Add(@"    Write-Host '[-] Owner  entfernen: " + this.DisplayName.PadRight(30) + " aus " + this.DisplayName + "'");
                            Global.TeamsPs1.Add(@"    Remove-UnifiedGroupLinks -Identity " + this.TeamId + " -LinkType Owner -Links '" + istOwner + "' -Confirm:$confirm"); // -Confirm:$false
                        }
                    }
                }
            }
        }
        /// <summary>
        /// Anrechnungen
        /// </summary>
        /// <param name="lehrers"></param>
        /// <param name="untisAnrechnung"></param>
        public Team(Lehrers lehrers, string untisAnrechnung)
        {
            this.Members = new List<string>();
            this.Owners = new List<string>();
            this.DisplayName = untisAnrechnung.Replace("--", "-")
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
                    if (anrechnung.Beschr == untisAnrechnung)
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