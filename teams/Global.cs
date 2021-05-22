using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.IO;

namespace teams
{
    public static class Global
    {
        public const string ConnectionStringUntis = @"Provider = Microsoft.Jet.OLEDB.4.0; Data Source=M:\\Data\\gpUntis.mdb;";
        public const string ConnectionStringAtlantis = @"Dsn=Atlantis9;uid=DBA";

        public static List<string> TeamsPs1 { get; set; }
        public static string TeamsPs { get; set; }
        public static string GruppenMemberPs { get; internal set; }

        public static string SafeGetString(OleDbDataReader reader, int colIndex)
        {
            if (!reader.IsDBNull(colIndex))
                return reader.GetString(colIndex);
            return string.Empty;
        }

        internal static string GruppenAuslesen(int anzahlTeamsIst)
        {
            return @"

if ((([System.Io.fileinfo]'C:\users\bm\Documents\GruppenOwnerMembers.csv').LastWriteTime.Date -ge [datetime]::Today)){     
    
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

        internal static string Ende()
        {
            return
                @"     
}else{
    
    Write-Host 'Da die GruppenOwnerMember.csv nicht von heute ist, wird sie nun erstellt. ...'
    
    Write-Host -ForegroundColor Green 'Alle Office 365-Gruppen werden geladen ...'
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
}
";
        }
    }
}