using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace teams
{
    class Program
    {
        private static IEnumerable<object> owner;

        static void Main(string[] args)
        {
            List<string> aktSj = new List<string>();
            
            DateTime datumMontagDerKalenderwoche = new DateTime(2020, 08, 10); //GetMondayDateOfWeek(kalenderwoche, DateTime.Now.Year);

            aktSj.Add((DateTime.Now.Month >= 8 ? DateTime.Now.Year : DateTime.Now.Year - 1).ToString());
            aktSj.Add((DateTime.Now.Month >= 8 ? DateTime.Now.Year + 1 : DateTime.Now.Year).ToString());

            const string connectionString = @"Provider = Microsoft.Jet.OLEDB.4.0; Data Source=M:\\Data\\gpUntis.mdb;";

            Periodes periodes = new Periodes(aktSj[0] + aktSj[1], connectionString);
            var periode = periodes.Count;

            Lehrers lehrers = new Lehrers(aktSj[0] + aktSj[1], connectionString, periodes);
            Klasses klasses = new Klasses(aktSj[0] + aktSj[1], lehrers, connectionString, periode);
            Schuelers schuelers = new Schuelers(klasses, aktSj[0] + aktSj[1]);
            Fachs fachs = new Fachs(aktSj[0] + aktSj[1], connectionString);
            Raums raums = new Raums(aktSj[0] + aktSj[1], connectionString, periode);
            Unterrichtsgruppes unterrichtsgruppes = new Unterrichtsgruppes(aktSj[0] + aktSj[1], connectionString);
            Unterrichts unterrichts = new Unterrichts(aktSj[0] + aktSj[1], datumMontagDerKalenderwoche, connectionString, periode, klasses, lehrers, fachs, raums,       unterrichtsgruppes);

            Teams teams = new Teams();

            Global.Streamwriter = new StreamWriter(@"C:\Users\bm\Documents\Teams.ps1");
            
            Global.Streamwriter.WriteLine("<# Connect to Exchange Online " + DateTime.Now.ToShortDateString() + " " +DateTime.Now.ToShortTimeString() + " #>");
            Global.Streamwriter.WriteLine("# $cred = Get-Credential");
            Global.Streamwriter.WriteLine("# $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $cred -Authentication Basic -AllowRedirection");
            Global.Streamwriter.WriteLine("# Import-PSSession $session");
            Global.Streamwriter.WriteLine("Connect-AzureAD -Credential $cred");            
            Global.Streamwriter.WriteLine("Connect-MicrosoftTeams -Credential $cred");
            Global.Streamwriter.WriteLine("");
            Global.Streamwriter.WriteLine("");
            
            Global.Streamwriter.WriteLine("$teamColl=Get-Team");
            Global.Streamwriter.WriteLine("$results = foreach($team in $teamColl)");
            Global.Streamwriter.WriteLine("{  ");
            Global.Streamwriter.WriteLine("Write-Host -ForegroundColor Magenta 'Getting all the owners from Team: ' $team.DisplayName  ");
            Global.Streamwriter.WriteLine("$ownerColl = Get-TeamUser -GroupId $team.GroupId -Role Owner  ");
            Global.Streamwriter.WriteLine("foreach($owner in $ownerColl)  ");
            Global.Streamwriter.WriteLine("{  ");
            Global.Streamwriter.WriteLine(" Write-Host -ForegroundColor Yellow 'User ID: ' $owner.UserId '   User: ' $owner.User  '   Name: ' $owner.Name  ");
            Global.Streamwriter.WriteLine("  [pscustomobject]@{");
            Global.Streamwriter.WriteLine("   Team = $team.GroupId");
            Global.Streamwriter.WriteLine("   User = $owner.User");
            Global.Streamwriter.WriteLine("   Role = 'Owner'");
            Global.Streamwriter.WriteLine("  }");
            Global.Streamwriter.WriteLine(" }");
            Global.Streamwriter.WriteLine(" $members = Get-TeamUser -GroupId $team.GroupId -Role Member  ");
            Global.Streamwriter.WriteLine(" foreach($member in $members)  ");
            Global.Streamwriter.WriteLine(" {  ");
            Global.Streamwriter.WriteLine(" Write-Host -ForegroundColor Yellow 'User ID: ' $member.UserId '   User: ' $member.User  '   Name: ' $member.Name  ");
            Global.Streamwriter.WriteLine("  [pscustomobject]@{");
            Global.Streamwriter.WriteLine("   Team = $team.GroupId");
            Global.Streamwriter.WriteLine("   User = $member.User");
            Global.Streamwriter.WriteLine("   Role = 'Member'");
            Global.Streamwriter.WriteLine("  }");
            Global.Streamwriter.WriteLine(" }  ");
            Global.Streamwriter.WriteLine("}  ");
            Global.Streamwriter.WriteLine(@"$results | Export-CSV -Path C:\users\bm\Documents\Teamsmembers.csv -NoTypeInformation");
            Global.Streamwriter.WriteLine(@"start notepad++ C:\users\bm\Documents\Teamsmembers.csv");
            
            // Kollegium

            List<string> owner = new List<string>() { (from l in lehrers where l.Kürzel == "BM" select l.Mail).FirstOrDefault() };
            List<string> member = (from l in lehrers where l.Mail != null where l.Mail.Contains("@") select l.Mail).ToList();
            member.Add("sonja.schwindowski@berufskolleg-borken.de");
            member.Add("simone.weiand@berufskolleg-borken.de");
            member.Add("evelyn.rietschel@berufskolleg-borken.de");
            member.Add("verena.baumeister@berufskolleg-borken.de");
            member.Add("sina.milewski@berufskolleg-borken.de");
            member.Add("sonja.berdychowski@berufskolleg-borken.de");
            member.Add("martina.domke@berufskolleg-borken.de");
            member.Add("ursula.moritz@berufskolleg-borken.de");
            NeueGruppe("Kollegium", owner, member, false, true);

            // Erweiterte Schulleitung
            
            //owner = (from l in lehrers where new List<string>() { "SH","AF","HE","KS","SC","SB","STH","TH","WE","SUE","HS","NX" }.Contains(l.Kürzel) select l.Mail).ToList();
            //member = new List<string>();
            //NeueGruppe("Erweiterte Schulleitung", owner, member, false, true);

            // Klassenleitungen

            owner = (from l in lehrers where new List<string>() { "BM" }.Contains(l.Kürzel) select l.Mail).ToList();
            member = (from k in (from kll in klasses where kll.Klassenleitungen != null where kll.Klassenleitungen.Count > 0 select kll).ToList() from kl in k.Klassenleitungen where kl != null where kl.Mail != null select kl.Mail).Distinct().ToList();
            NeueGruppe("Klassenleitungen", owner, member, false, false);

            // Lehrerrat

            //owner = (from l in lehrers where new List<string>() { "HC", "STA", "HR", "GU", "GV" }.Contains(l.Kürzel) select l.Mail).ToList();
            //member = new List<string>();
            //NeueGruppe("Erweiterte Schulleitung", owner, member, false, true);

            // Sekretariat

            //owner = (from l in lehrers where new List<string>() { "BM" }.Contains(l.Kürzel) select l.Mail).ToList();
            //member = (from l in lehrers where new List<string>() { "BM", "SUE" }.Contains(l.Kürzel) select l.Mail).ToList();
            //member.Add("sonja.schwindowski@berufskolleg-borken.de");
            //member.Add("simone.weiand@berufskolleg-borken.de");
            //member.Add("evelyn.rietschel@berufskolleg-borken.de");
            //member.Add("martina.domke@berufskolleg-borken.de");
            //NeueGruppe("Sekretariat", owner, member, false, true);

            // Beratung

            //owner = (from l in lehrers where new List<string>() { "BM" }.Contains(l.Kürzel) select l.Mail).ToList();
            //member = (from l in lehrers where new List<string>() { "TH", "WK", "STH", "STA", "LN", "HR" }.Contains(l.Kürzel) select l.Mail).ToList();
            //NeueGruppe("Beratung", owner, member, false, true);

            // Medienkoordination

            //owner = (from l in lehrers where new List<string>() { "BM" }.Contains(l.Kürzel) select l.Mail).ToList();
            //member = (from l in lehrers where new List<string>() { "KEL", "FIN" }.Contains(l.Kürzel) select l.Mail).ToList();
            //NeueGruppe("Medienkoordination", owner, member, false, false);
            
            // Schulsozialarbeit

            //owner = new List<string>();
            //owner.Add("sina.milewski@berufskolleg-borken.de");
            //owner.Add("sonja.berdychowski@berufskolleg-borken.de");
            //member = new List<string>();
            //NeueGruppe("Schulsozialarbeit", owner, member, false, false);
            
            // Klassengruppen

            foreach (var klasse in (from k in klasses where (k.Klassenleitungen != null && k.Klassenleitungen.Count > 0 && k.Klassenleitungen[0] != null) select k))
            {
                owner = (from l in lehrers where ((from u in unterrichts where u.KlasseKürzel == klasse.NameUntis select u.LehrerKürzel).ToList()).Contains(l.Kürzel) select l.Mail).ToList();
                member = (from s in schuelers where s.Klasse == klasse.NameUntis select s.Mail).ToList();
                NeueGruppe(klasse.NameUntis, owner, member, false, true);
            }
            
            Global.Streamwriter.Close();            
        }

        private static void NeueGruppe(string displayName, List<string> owner, List<string> member, bool inOutlookVerstecken, bool teamAnlegen)
        {
            Global.Streamwriter.WriteLine("");
            Global.Streamwriter.WriteLine(@"Write-Host('*********************************************************************************************')");
            Global.Streamwriter.WriteLine("");
            Global.Streamwriter.WriteLine(@"Write-Host('" + displayName + " anlegen ...')");
            Global.Streamwriter.WriteLine("");
            if (owner.Count > 0)
            {
                // Existenz prüfen

                Global.Streamwriter.WriteLine("if (-not ([string]::IsNullOrEmpty((Get-UnifiedGroup | Where-Object{($_.PrimarySmtpAddress -eq '" + displayName.Replace(' ', '-') + "@berufskolleg-borken.de')} | Select-Object Name).Name)))");
                Global.Streamwriter.WriteLine("{ Write-Host('"+ displayName +" existiert bereits.') }");
                Global.Streamwriter.WriteLine("else { New-UnifiedGroup -DisplayName '" + displayName + "' -EmailAddresses: '" + displayName.Replace(' ', '-') + "@berufskolleg-borken.de' -Notes '" + displayName + "' -Language 'de-DE' -AccessType Private -Owner '" + owner[0] + "' -AutoSubscribeNewMembers }");
                Global.Streamwriter.WriteLine("");                
            }
            else
            {
                Global.Streamwriter.WriteLine("if (-not ([string]::IsNullOrEmpty((Get-UnifiedGroup | Where-Object{($_.PrimarySmtpAddress -eq '" + displayName.Replace(' ', '-') + "@berufskolleg-borken.de')} | Select-Object Name).Name)))");
                Global.Streamwriter.WriteLine("{ Write-Host('" + displayName + " existiert bereits.') }");
                Global.Streamwriter.WriteLine("else { New-UnifiedGroup -DisplayName '" + displayName + "' -EmailAddresses: '" + displayName.Replace(' ', '-') + "@berufskolleg-borken.de' -Notes '" + displayName + "' -Language 'de-DE' -AccessType Private -AutoSubscribeNewMembers }");
                Global.Streamwriter.WriteLine("");
            }
            
            Global.Streamwriter.WriteLine("$teamsgroups = (Get-UnifiedGroup | Where-Object{($_.PrimarySmtpAddress -eq '" + displayName.Replace(' ', '-') + "@berufskolleg-borken.de')} | Select-Object Name).Name");
            
            Global.Streamwriter.WriteLine("Get-UnifiedGroup -Identity $teamsgroups | Format-List DisplayName,EmailAddresses,Notes,ManagedBy,AccessType");

            // in Outlook verstecken

            if (inOutlookVerstecken)
            {
                Global.Streamwriter.WriteLine(@"Write-Host('" + displayName + " in Outlook verstecken')");
                Global.Streamwriter.WriteLine("Set-UnifiedGroup -Identity $teamsgroups -HiddenFromExchangeClientsEnabled");
            }
            else
            {
                Global.Streamwriter.WriteLine("Set-UnifiedGroup -Identity $teamsgroups -HiddenFromExchangeClientsEnabled:$False");
                Global.Streamwriter.WriteLine("Set-UnifiedGroup -Identity $teamsgroups -HiddenFromAddressListsEnabled $false");
            }
            
            Global.Streamwriter.WriteLine("Get-UnifiedGroup -Identity $teamsgroups | ft DisplayName,HiddenFrom*");
            Global.Streamwriter.WriteLine("");

            // User

            Global.Streamwriter.WriteLine("");
            Global.Streamwriter.WriteLine(@"Write-Host('" + displayName + " Owner anlegen ...')");

            foreach (var o in owner)
            {                
                Global.Streamwriter.WriteLine(@"Add-UnifiedGroupLinks -Identity $teamsgroups -LinkType Owner -Links '" + o + "'");
            }

            Global.Streamwriter.WriteLine(@"Get-UnifiedGroup -Identity $teamsgroups | Get-UnifiedGroupLinks -LinkType Owner");

            Global.Streamwriter.WriteLine("");
            Global.Streamwriter.WriteLine(@"Write-Host('" + displayName + " Member anlegen ...')");

            foreach (var m in member)
            {             
                Global.Streamwriter.WriteLine(@"Add-UnifiedGroupLinks -Identity $teamsgroups -LinkType Member -Links '" + m + "'");
            }

            Global.Streamwriter.WriteLine(@"Get-UnifiedGroup -Identity $teamsgroups | Get-UnifiedGroupLinks -LinkType Member");

            if (teamAnlegen)
            {
                Global.Streamwriter.WriteLine("");
                Global.Streamwriter.WriteLine(@"Write-Host('" + displayName + " Team anlegen')");
                Global.Streamwriter.WriteLine("$adGruppe = Get-AzureADGroup -SearchString '" + displayName + "'");
                Global.Streamwriter.WriteLine("$group = New-Team -group $adGruppe.ObjectId");
            }
        }

        private static DateTime GetMondayDateOfWeek(int week, int year)
        {
            int i = CultureInfo.CurrentCulture.Calendar.GetWeekOfYear(new DateTime(year, 1, 1), CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);

            if (i == 1)
            {
                return CultureInfo.CurrentCulture.Calendar.AddDays(new DateTime(year, 1, 1), ((week - 1) * 7 - GetDayCountFromMonday(CultureInfo.CurrentCulture.Calendar.GetDayOfWeek(new DateTime(year, 1, 1))) + 1));
            }
            else
            {
                int x = Convert.ToInt32(CultureInfo.CurrentCulture.Calendar.GetDayOfWeek(new DateTime(year, 1, 1)));
                return CultureInfo.CurrentCulture.Calendar.AddDays(new DateTime(year, 1, 1), ((week - 1) * 7 + (7 - GetDayCountFromMonday(CultureInfo.CurrentCulture.Calendar.GetDayOfWeek(new DateTime(year, 1, 1)))) + 1));
            }
        }

        private static int GetDayCountFromMonday(DayOfWeek dayOfWeek)
        {
            switch (dayOfWeek)
            {
                case DayOfWeek.Monday:
                    return 1;
                case DayOfWeek.Tuesday:
                    return 2;
                case DayOfWeek.Wednesday:
                    return 3;
                case DayOfWeek.Thursday:
                    return 4;
                case DayOfWeek.Friday:
                    return 5;
                case DayOfWeek.Saturday:
                    return 6;
                default:
                    //Sunday
                    return 7;
            }
        }

        public int GetCalendarWeek(DateTime date)
        {
            CultureInfo currentCulture = CultureInfo.CurrentCulture;

            Calendar calendar = currentCulture.Calendar;

            int calendarWeek = calendar.GetWeekOfYear(date,
               currentCulture.DateTimeFormat.CalendarWeekRule,
               currentCulture.DateTimeFormat.FirstDayOfWeek);

            // Überprüfen, ob eine Kalenderwoche größer als 52
            // ermittelt wurde und ob die Kalenderwoche des Datums
            // in einer Woche 2 ergibt: In diesem Fall hat
            // GetWeekOfYear die Kalenderwoche nicht nach ISO 8601 
            // berechnet (Montag, der 31.12.2007 wird z. B.
            // fälschlicherweise als KW 53 berechnet). 
            // Die Kalenderwoche wird dann auf 1 gesetzt
            if (calendarWeek > 52)
            {
                date = date.AddDays(7);
                int testCalendarWeek = calendar.GetWeekOfYear(date,
                   currentCulture.DateTimeFormat.CalendarWeekRule,
                   currentCulture.DateTimeFormat.FirstDayOfWeek);
                if (testCalendarWeek == 2)
                    calendarWeek = 1;
            }
            return calendarWeek;
        }
    }
}
