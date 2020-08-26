﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;

namespace teams
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                List<string> aktSj = new List<string>();

                DateTime datumMontagDerKalenderwoche = new DateTime(2020, 08, 10); //GetMondayDateOfWeek(kalenderwoche, DateTime.Now.Year);
                Global.TeamsPs = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal) + @"\\Teams.ps1";
                Global.GruppenMemberPs = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal) + @"\\GruppenOwnerMembers.csv";

                Global.TeamsPs1 = new List<string>();
                aktSj.Add((DateTime.Now.Month >= 8 ? DateTime.Now.Year : DateTime.Now.Year - 1).ToString());
                aktSj.Add((DateTime.Now.Month >= 8 ? DateTime.Now.Year + 1 : DateTime.Now.Year).ToString());

                const string connectionString = @"Provider = Microsoft.Jet.OLEDB.4.0; Data Source=M:\\Data\\gpUntis.mdb;";

                Periodes periodes = new Periodes(aktSj[0] + aktSj[1], connectionString);
                var periode = periodes.Count;

                Lehrers lehrers = new Lehrers(aktSj[0] + aktSj[1], connectionString, periodes);
                Klasses klasses = new Klasses(aktSj[0] + aktSj[1], lehrers, connectionString, periode);

                if (!File.Exists(Global.TeamsPs))
                {
                    File.Create(Global.TeamsPs);
                    File.Create(Global.GruppenMemberPs);
                }

                try
                {
                    File.WriteAllText(Global.TeamsPs, Global.Auth());
                    File.AppendAllText(Global.TeamsPs, Global.GruppenAuslesen());
                }
                catch (Exception ex)
                {
                    throw new Exception(" Die Datei Teams.ps1 ist von einem anderen Prozess gesperrt.\n\n" + ex);                    
                }

                Teams teamsIst = new Teams(Global.GruppenMemberPs, klasses);

                Schuelers schuelers = new Schuelers(klasses, aktSj[0] + aktSj[1]);
                Fachs fachs = new Fachs(aktSj[0] + aktSj[1], connectionString);
                Raums raums = new Raums(aktSj[0] + aktSj[1], connectionString, periode);
                Unterrichtsgruppes unterrichtsgruppes = new Unterrichtsgruppes(aktSj[0] + aktSj[1], connectionString);
                Unterrichts unterrichts = new Unterrichts(aktSj[0] + aktSj[1], datumMontagDerKalenderwoche, connectionString, periode, klasses, lehrers, fachs, raums, unterrichtsgruppes);

                Teams klassenTeamsSoll = new Teams(klasses, lehrers, schuelers, unterrichts);

                //klassenTeamsSoll.FehlendeKlassenAnlegen(teamsIst);

                teamsIst.DoppelteKlassenFinden();

                klassenTeamsSoll.OwnerUndMemberAnlegen(teamsIst);

                teamsIst.OwnerUndMemberLöschen(klassenTeamsSoll);

                
                Global.TeamsPs1.Add("Write-Host 'Ende der Verarbeitung'");
                File.AppendAllLines(Global.TeamsPs, Global.TeamsPs1);

                //Process.Start("powershell_ise.exe", Global.TeamsPs);

                Console.WriteLine("Verarbeitung beendet");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Console.ReadKey();
            }
            finally
            {
                Environment.Exit(0);                
            }
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
