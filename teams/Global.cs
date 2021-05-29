using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.IO;
using System.Text;
using System.Threading;

namespace teams
{
    public static class Global
    {
        public static List<string> TeamsPs1 { get; set; }
        public static string TeamsPs { get; set; }
        public static string GruppenMemberPs { get; internal set; }

        public const string ConnectionStringUntis = @"Provider = Microsoft.Jet.OLEDB.4.0; Data Source=M:\\Data\\gpUntis.mdb;";
        public const string ConnectionStringAtlantis = @"Dsn=Atlantis9;uid=DBA";
        public static List<string> AktSj = new List<string>() {
            (DateTime.Now.Month >= 8 ? DateTime.Now.Year : DateTime.Now.Year - 1).ToString(),
            (DateTime.Now.Month >= 8 ? DateTime.Now.Year + 1 : DateTime.Now.Year).ToString()
        };

        internal static void Initialize()
        {
            Console.WriteLine("      teams.exe | Published under the terms of GPLv3 | Stefan Bäumer " + DateTime.Now.Year + " | Version 20210525");
            Console.WriteLine("===================================================================================================");
            Console.WriteLine(" *teams.exe* erstellt eine Powershell-Datei mit entsprechenden Befehlen zum Abgleich mit Office365.");
            Console.WriteLine(" In einem ersten Durchauf wird eine CSV-Datei mit den vorhandenen Teams und Verteilergruooen ausge-");
            Console.WriteLine(" lesen.  In einem zweiten Durchlauf werden die Teams, Verteilergruppen, Owner & Member abgeglichen.");
            Console.WriteLine(" Führen Sie die Powershell als Administrator in der ISE aus.");
            Console.WriteLine("===================================================================================================");

            Global.TeamsPs = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal) + @"\\Teams.ps1";
            Global.GruppenMemberPs = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal) + @"\\GruppenOwnerMembers.csv";

            File.WriteAllText(Global.TeamsPs, @"# Skript zum tagesaktuellen Abgleich der Klassen aus Atlantis / Untis mit Office 365
", Encoding.UTF8);
            File.AppendAllLines(Global.TeamsPs, new List<string>() { "# " + DateTime.Now.Date.ToShortDateString() + " " + DateTime.Now.ToShortTimeString() + "" }, Encoding.UTF8);
            Global.TeamsPs1 = new List<string>();

            try
            {
                File.AppendAllLines(Global.TeamsPs, new List<string>() { "", "$confirm = $true" }, Encoding.UTF8);
                File.AppendAllText(Global.TeamsPs, Global.Auth(), Encoding.UTF8);
            }
            catch (Exception ex)
            {
                throw new Exception(" Die Datei Teams.ps1 ist von einem anderen Prozess gesperrt.\n\n" + ex);
            }

            CheckGruppenOwnersMembers();
        }

        /// <summary>
        /// Wenn die GruppenOwnersMembers nicht existiert oder nicht von heute ist, wird der Powershell-Befhel zum Auslesen der Gruppen generiert ...
        /// </summary>
        private static void CheckGruppenOwnersMembers()
        {
            var gruppenOwnersMembers = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "GruppenOwnerMembers.csv");

            if (File.GetLastWriteTime(gruppenOwnersMembers).Date < DateTime.Now.Date || !File.Exists(gruppenOwnersMembers))
            {
                if (!File.Exists(gruppenOwnersMembers))
                {
                    File.Create(gruppenOwnersMembers);
                }
                File.AppendAllText(Global.TeamsPs, Global.GruppenLesen(), Encoding.UTF8);
                Console.WriteLine("Verarbeitung beendet.");
                Console.WriteLine("Öffnen Sie jetzt die ISE als Administrator. Öffnen Sie in der ISE die Datei " + Global.TeamsPs + ". ");
                Thread.Sleep(10000); //will sleep for 5 sec
                Environment.Exit(0);
            }
        }

        public static string SafeGetString(OleDbDataReader reader, int colIndex)
        {
            if (!reader.IsDBNull(colIndex))
                return reader.GetString(colIndex);
            return string.Empty;
        }

        internal static string GruppenAuslesen(int anzahlTeamsIst)
        {
            return @"

  
    
    Write-Host '|'
    Write-Host '| Da die GruppenOwnerMember.csv heute zuletzt aktualisiert wurde, wird ein Abgleich aller Gruppen mit Ownern und Membern gemacht ...'
    Write-Host '|'

";
        }

        internal static string Auth()
        {
            return @"
$testSession = Get-PSSession
if(-not($testSession))
{
    Write-Warning '$targetComputer : Sie sind Nicht angemeldet.'
    $cred = Get-Credential
    $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $cred -Authentication Basic -AllowRedirection
    Import-PSSession $session
    Connect-AzureAD -Credential $cred
    Connect-MicrosoftTeams -Credential $cred
}
else
{
    Write-Host '$targetComputer: Sie sind angemeldet.'    
}";


        }

        internal static string GruppenLesen()
        {
        return
@"     
# Da die GruppenOwnerMember.csv nicht von heute ist, wird sie nun erstellt. ...
# Anschließend muss Teams.exe erneut gestartet werden.    

Write-Host -ForegroundColor Green 'Alle Office 365-Gruppen und Verteilergruppen werden geladen ...'
$Groups = Get-UnifiedGroup -ResultSize Unlimited  | Sort-Object DisplayName

# Process Groups
$GroupsCSV = @()

Write-Host -ForegroundColor Green 'Processing Groups'

$results = foreach ($Group in $Groups)
{
    Write-Host -ForegroundColor Magenta 'Hole alle  Owner der Gruppe ' $Group.DisplayName  '('$Group.Identity')' ...
    $Owners = Get-UnifiedGroupLinks -Identity $Group.Identity -LinkType Owners -ResultSize Unlimited
    
    foreach ($Owner in $Owners)
    {
         
            [pscustomobject]@{
            GroupId = $Group.Identity
            GroupDisplayName = $Group.DisplayName
            User = $Owner.PrimarySmtpAddress
            Role = 'Owner'
            Type = 'O365'
        }
    }
 
        Write-Host -ForegroundColor Magenta 'Hole alle Member der Gruppe ' $Group.DisplayName  '('$Group.Identity')' ...
        $Members = Get-UnifiedGroupLinks -Identity $Group.Identity -LinkType Members -ResultSize Unlimited
        $MembersSMTP=@()
    
    foreach ($Member in $Members)
    {
        [pscustomobject]@{
            GroupId = $Group.Identity
            GroupDisplayName = $Group.DisplayName
            User = $Member.PrimarySmtpAddress
            Role = 'Member'
            Type = 'O365'
        }        
    }        
}

Write-Host -ForegroundColor Green 'Alle Verteilergruppen werden geladen'
$Groups = Get-DistributionGroup -ResultSize Unlimited | Sort-Object DisplayName
       

# Process Groups
    
Write-Host -ForegroundColor Green 'Processing Groups'

$resultsV = foreach ($Group in $Groups)
{
    Write-Host -ForegroundColor Magenta 'Hole alle Member der Verteilergruppe ' $Group.DisplayName  '('$Group.Identity')' ...
        $Members = Get-DistributionGroupMember -Identity $Group.Identity -ResultSize Unlimited
        $MembersSMTP=@()
    
    foreach ($Member in $Members)
    {
        [pscustomobject]@{
            GroupId = $Group.Identity
            GroupDisplayName = $Group.DisplayName
            User = $Member.PrimarySmtpAddress
            Role = 'Member'
            Type = 'Distribution'
        }        
    }        
}

$results = $results + $resultsV

# Export to CSV
Write-Host -ForegroundColor Green 'GruppenOwnerMembers.csv wird geschrieben. Nun kann Teams.exe erneut gestartet werden.'
$results | Export-Csv -NoTypeInformation -Path C:\users\bm\Documents\GruppenOwnerMembers.csv -Encoding UTF8 -Delimiter '|'
start notepad++ C:\users\bm\Documents\GruppenOwnerMembers.csv    
start-process -FilePath 'U:\Source\Repos\teams\teams\bin\Debug\teams.exe'

";
        }

        internal static void WriteLine(string v, int count)
        {
            Console.WriteLine((v + " " + ".".PadRight(count / 150, '.')).PadRight(48, '.') + (" " + count).ToString().PadLeft(4), '.');
            File.AppendAllLines(Global.TeamsPs, new List<string>() { "# Anzahl " + v + " : " + count + "" }, Encoding.UTF8);
        }
    }
}