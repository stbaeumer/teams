# Get all the teams from tenant  
$teamColl=Get-Team -DisplayName "GW18A" 
 
# Loop through the teams  
$results = foreach($team in $teamColl)  
{  
    Write-Host -ForegroundColor Magenta "Getting all the owners from Team: " $team.DisplayName  
 
    # Get the team owners  
    
    $ownerColl = Get-TeamUser -GroupId $team.GroupId -Role Owner  
 
    #Loop through the owners  
    
    foreach($owner in $ownerColl)  
    {  
        Write-Host -ForegroundColor Yellow "User ID: " $owner.UserId "   User: " $owner.User  "   Name: " $owner.Name  

        [pscustomobject]@{
            Team = $team.GroupId
            User = $owner.User
            Role = 'Owner'
        }
    }
    
    # Get the team members  
    
    $members = Get-TeamUser -GroupId $team.GroupId -Role Member  
      
    foreach($member in $members)  
    {  
        Write-Host -ForegroundColor Yellow "User ID: " $member.UserId "   User: " $member.User  "   Name: " $member.Name  

        [pscustomobject]@{
            Team = $team.GroupId
            User = $member.User
            Role = 'Member'
        }
    }      
}  

$results | Export-CSV -Path C:\users\bm\Documents\Teamsmembers.csv -NoTypeInformation
start notepad++ C:\users\bm\Documents\Teamsmembers.csv
