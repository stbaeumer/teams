using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

namespace teams
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Global.Initialize();

                Periodes periodes = new Periodes();

                var periode = (from p in periodes where p.Bis >= DateTime.Now.Date where DateTime.Now.Date >= p.Von select p.IdUntis).FirstOrDefault();

                Lehrers lehrers = new Lehrers(periode);
                Klasses klasses = new Klasses(periode, lehrers);
                Schuelers schuelers = new Schuelers(klasses);
                Fachs fachs = new Fachs();
                Raums raums = new Raums(periode);
                Unterrichtsgruppes unterrichtsgruppes = new Unterrichtsgruppes();
                Unterrichts unterrichts = new Unterrichts(periode, klasses, lehrers, fachs, raums, unterrichtsgruppes);

                Teams teamsIst = new Teams(Global.GruppenMemberPs, klasses);
                
                Teams klassenteamsSoll = new Teams(klasses, lehrers, schuelers, unterrichts);
                
                teamsIst.SyncTeam(false, true, lehrers, new Team(klassenteamsSoll, "Kurs-20-Abiturienten"));
                teamsIst.SyncTeam(true, false, lehrers, new Team(klassenteamsSoll, "Gym13LuL"));
                teamsIst.SyncTeam(true, false, lehrers, new Team(klassenteamsSoll, "BCAbschlussklassenLuL"));
                teamsIst.SyncTeam(true, false, lehrers, new Team(klassenteamsSoll, "BlaueBriefe"));
                teamsIst.SyncTeam(false, true, lehrers, new Team(lehrers)); // Kollegium

                teamsIst.SyncTeams(true, false, lehrers, new Teams(klasses, lehrers)); // Klassenleitung
                teamsIst.SyncTeams(true, false, lehrers, new Teams(lehrers, "Lehrerinnen"));
                teamsIst.SyncTeams(true, false, lehrers, new Teams(lehrers, "Vollzeitkraefte"));
                teamsIst.SyncTeams(true, false, lehrers, new Teams(lehrers, "Teilzeitkraefte"));
                teamsIst.SyncTeams(true, false, lehrers, new Teams(klassenteamsSoll));
                
                // Teams aus den Untis-Anrechnungen

                foreach (var untisAnrechnung in GetUntisAnrechnungen(lehrers))
                {
                    teamsIst.SyncTeam(true, false, lehrers, new Team(lehrers, untisAnrechnung));
                }

                if (!File.Exists(Global.TeamsPs))
                {
                    File.Create(Global.TeamsPs);
                    File.Create(Global.GruppenMemberPs);
                }

                Global.TeamsPs1.Add("    Write-Host 'Ende der Verarbeitung'");
                File.AppendAllLines(Global.TeamsPs, Global.TeamsPs1, Encoding.UTF8);
                Console.ReadKey();
                Environment.Exit(0);

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Console.ReadKey();
            }
            finally
            {

            }
        }

        private static List<string> GetUntisAnrechnungen(Lehrers lehrers)
        {
            var untisanrechnungen =new List<string>();
            untisanrechnungen.AddRange((from l in lehrers from a in l.Anrechnungen select a.Beschr).ToList().Distinct());

            foreach (var an in (from l in lehrers from a in l.Anrechnungen select a.TextGekürzt).Distinct().ToList())
            {
                if (!untisanrechnungen.Contains(an))
                {
                    untisanrechnungen.Add(an);
                }
            }

            return untisanrechnungen;
        }

        //private static void NeueGruppe(string displayName, List<string> owner, List<string> member, bool inOutlookVerstecken, bool teamAnlegen)
        //{
        //    File.AppendAllText(Global.TeamsPs1,"");
        //    File.AppendAllText(Global.TeamsPs1,@"Write-Host('*********************************************************************************************')");
        //    File.AppendAllText(Global.TeamsPs1,"");
        //    File.AppendAllText(Global.TeamsPs1,@"Write-Host('" + displayName + " anlegen ...')");
        //    File.AppendAllText(Global.TeamsPs1,"");
        //    if (owner.Count > 0)
        //    {
        //        // Existenz prüfen

        //        File.AppendAllText(Global.TeamsPs1,"if (-not ([string]::IsNullOrEmpty((Get-UnifiedGroup | Where-Object{($_.PrimarySmtpAddress -eq '" + displayName.Replace(' ', '-') + "@berufskolleg-borken.de')} | Select-Object Name).Name)))");
        //        File.AppendAllText(Global.TeamsPs1,"{ Write-Host('"+ displayName +" existiert bereits.') }");
        //        File.AppendAllText(Global.TeamsPs1,"else { New-UnifiedGroup -DisplayName '" + displayName + "' -EmailAddresses: '" + displayName.Replace(' ', '-') + "@berufskolleg-borken.de' -Notes '" + displayName + "' -Language 'de-DE' -AccessType Private -Owner '" + owner[0] + "' -AutoSubscribeNewMembers }");
        //        File.AppendAllText(Global.TeamsPs1,"");                
        //    }
        //    else
        //    {
        //        File.AppendAllText(Global.TeamsPs1,"if (-not ([string]::IsNullOrEmpty((Get-UnifiedGroup | Where-Object{($_.PrimarySmtpAddress -eq '" + displayName.Replace(' ', '-') + "@berufskolleg-borken.de')} | Select-Object Name).Name)))");
        //        File.AppendAllText(Global.TeamsPs1,"{ Write-Host('" + displayName + " existiert bereits.') }");
        //        File.AppendAllText(Global.TeamsPs1,"else { New-UnifiedGroup -DisplayName '" + displayName + "' -EmailAddresses: '" + displayName.Replace(' ', '-') + "@berufskolleg-borken.de' -Notes '" + displayName + "' -Language 'de-DE' -AccessType Private -AutoSubscribeNewMembers }");
        //        File.AppendAllText(Global.TeamsPs1,"");
        //    }

        //    File.AppendAllText(Global.TeamsPs1,"$teamsgroups = (Get-UnifiedGroup | Where-Object{($_.PrimarySmtpAddress -eq '" + displayName.Replace(' ', '-') + "@berufskolleg-borken.de')} | Select-Object Name).Name");

        //    File.AppendAllText(Global.TeamsPs1,"Get-UnifiedGroup -Identity $teamsgroups | Format-List DisplayName,EmailAddresses,Notes,ManagedBy,AccessType");

        //    // in Outlook verstecken

        //    if (inOutlookVerstecken)
        //    {
        //        File.AppendAllText(Global.TeamsPs1,@"Write-Host('" + displayName + " in Outlook verstecken')");
        //        File.AppendAllText(Global.TeamsPs1,"Set-UnifiedGroup -Identity $teamsgroups -HiddenFromExchangeClientsEnabled");
        //    }
        //    else
        //    {
        //        File.AppendAllText(Global.TeamsPs1,"Set-UnifiedGroup -Identity $teamsgroups -HiddenFromExchangeClientsEnabled:$False");
        //        File.AppendAllText(Global.TeamsPs1,"Set-UnifiedGroup -Identity $teamsgroups -HiddenFromAddressListsEnabled $false");
        //    }

        //    File.AppendAllText(Global.TeamsPs1,"Get-UnifiedGroup -Identity $teamsgroups | ft DisplayName,HiddenFrom*");
        //    File.AppendAllText(Global.TeamsPs1,"");

        //    // User

        //    File.AppendAllText(Global.TeamsPs1,"");
        //    File.AppendAllText(Global.TeamsPs1,@"Write-Host('" + displayName + " Owner anlegen ...')");

        //    foreach (var o in owner)
        //    {                
        //        File.AppendAllText(Global.TeamsPs1,@"Add-UnifiedGroupLinks -Identity $teamsgroups -LinkType Owner -Links '" + o + "'");
        //    }

        //    File.AppendAllText(Global.TeamsPs1,@"Get-UnifiedGroup -Identity $teamsgroups | Get-UnifiedGroupLinks -LinkType Owner");

        //    File.AppendAllText(Global.TeamsPs1,"");
        //    File.AppendAllText(Global.TeamsPs1,@"Write-Host('" + displayName + " Member anlegen ...')");

        //    foreach (var m in member)
        //    {             
        //        File.AppendAllText(Global.TeamsPs1,@"Add-UnifiedGroupLinks -Identity $teamsgroups -LinkType Member -Links '" + m + "'");
        //    }

        //    File.AppendAllText(Global.TeamsPs1,@"Get-UnifiedGroup -Identity $teamsgroups | Get-UnifiedGroupLinks -LinkType Member");

        //    if (teamAnlegen)
        //    {
        //        File.AppendAllText(Global.TeamsPs1,"");
        //        File.AppendAllText(Global.TeamsPs1,@"Write-Host('" + displayName + " Team anlegen')");
        //        File.AppendAllText(Global.TeamsPs1,"$adGruppe = Get-AzureADGroup -SearchString '" + displayName + "'");
        //        File.AppendAllText(Global.TeamsPs1,"$group = New-Team -group $adGruppe.ObjectId");
        //    }
        //}

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
