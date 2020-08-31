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

        internal static string GruppenAuslesen()
        {
            return @"

if ( ([System.Io.fileinfo]'C:\users\bm\Documents\GruppenOwnerMembers.csv').LastWriteTime.Date -ge [datetime]::Today ){     
# if ( $FALSE ){     
     'Die Datei C:\users\bm\Documents\GruppenOwnerMembers.csv muss nicht aktualisiert werden.'
}else{

    Write-Host -ForegroundColor Green 'Loading all Office 365 Groups'
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
            }        
        }        
    }

    # Export to CSV
    Write-Host -ForegroundColor Green 'GruppenOwnerMembers.csv wird geschrieben'
    $results | Export-Csv -NoTypeInformation -Path C:\users\bm\Documents\GruppenOwnerMembers.csv
    start notepad++ C:\users\bm\Documents\GruppenOwnerMembers.csv    
}

";
        }

        internal static string Auth()
        {
            return @"

$testSession = Get-PSSession
if(-not($testSession))
{
    Write-Warning '$targetComputer : Nicht angemeldet!'
    $cred = Get-Credential
    $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $cred -Authentication Basic -AllowRedirection
    Import-PSSession $session
    Connect-AzureAD -Credential $cred
    Connect-MicrosoftTeams -Credential $cred
}
else
{
    Write-Host '$targetComputer: Angemeldet!'    
}";


        }
    }
}