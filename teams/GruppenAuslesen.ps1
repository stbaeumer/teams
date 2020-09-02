#
.SYNOPSIS
  Example script to export Office 365 Groups with Members and Owners to CSV
.DESCRIPTION
  This is an example script from Practical365.com to demonstrate exporting Office 365 Group / Team members and owners to CSV files
.PARAMETER $CSVFilename
  The path the the CSV file
.INPUTS
  None
.OUTPUTS
  CSV File
.NOTES
  Version:        1.0
  Author:         Steve Goodman
  Creation Date:  20190113
  Purpose/Change: Initial version
  
.EXAMPLE
  Export-O365Groups -CSVFilename .\Groups.CSV
#>

Write-Host -ForegroundColor Green "Loading all Office 365 Groups"
$Groups = Get-UnifiedGroup -ResultSize Unlimited -Identity "GW18A"

# Process Groups
$GroupsCSV = @()

Write-Host -ForegroundColor Green "Processing Groups"

$results = foreach ($Group in $Groups)
{
    Write-Host -ForegroundColor Magenta "Hole alle Owner der Gruppe " $Group.DisplayName  "("$Group.Identity")" ...
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
 
     Write-Host -ForegroundColor Magenta "Hole alle Member der Gruppe " $Group.DisplayName  "("$Group.Identity")" ...
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
Write-Host -ForegroundColor Green "`nCreating and exporting CSV file"
$results | Export-Csv -NoTypeInformation -Path C:\users\bm\Documents\GruppenOwnerMembers.csv
start notepad++ C:\users\bm\Documents\GruppenOwnerMembers.csv