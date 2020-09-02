# Connect to Exchange Online #>

#$cred = Get-Credential
#$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $cred -Authentication Basic -AllowRedirection
#Import-PSSession $session

# Office 365-Gruppe mit Owner anlegen #>

New-UnifiedGroup -DisplayName "Kollegium" -EmailAddresses: "SMTP:Kollegium@berufskolleg-borken.de" -Notes "Kollegium" -Language "de-DE" -AccessType Private -Owner "stefan.baeumer@berufskolleg-borken.de" -AutoSubscribeNewMembers

# Sichtbarkeit von Teams-basierten Gruppen in Outlook (direkter Mailverteiler der Gruppe wird sichtbar)#>

# include the new group in the Outlook client’s Groups navigation pane #>
 
Set-UnifiedGroup -Identity "Kollegium" -HiddenFromExchangeClientsEnabled:$False

# include the new group in the Exchange Online Outlook Address Book #>

Set-UnifiedGroup -Identity "Kollegium" -HiddenFromAddressListsEnabled $false
Get-UnifiedGroup -Identity "Kollegium" | ft DisplayName,HiddenFrom*

## Weitere Owner anlegen

# Add-UnifiedGroupLinks -Identity "Testgruppe7" -LinkType Members -Links "stefan.baeumer@berufskolleg-borken.de"
# Add-UnifiedGroupLinks -Identity "Testgruppe7" -LinkType Owner -Links "stefan.baeumer@berufskolleg-borken.de"

# Member anlegen #>

# Add-UnifiedGroupLinks -Identity "Kollegium" -LinkType Members -Links "stefan.baeumer@berufskolleg-borken.de"

# Aus der O365-Gruppe wird ein Team erzeugt #>
# https://www.tecklyfe.com/how-to-create-a-microsoft-team-from-an-office-365-group-using-powershell/ #>

#$adGruppe = Get-AzureADGroup -SearchString "Testgruppe7"
#$adGruppe.ObjectId
#$group = New-Team -group $adGruppe.ObjectId -Template EDU_Class







# Remove-PSSession $Session







#Install-Module MicrosoftTeams
#Connect-AzureAD
#Connect-MicrosoftTeams


# Owner anlegen #>

#Add-UnifiedGroupLinks -Identity "Testgruppe" -LinkType Members -Links "stefan.baeumer@berufskolleg-borken.de"
#Add-UnifiedGroupLinks -Identity "Testgruppe" -LinkType Owner -Links "stefan.baeumer@berufskolleg-borken.de"






#$group = New-Team -DisplayName "K3" -Visibility "private" -Description "Dies ist ein neues Team"
Add-TeamUser -GroupId $group.GroupId -User "stefan.baeumer@berufskolleg-borken.de" -Role "owner"
Add-TeamUser -GroupId $group.GroupId -User "evelyn.rietschel@berufskolleg-borken.de" -Role "member"

 New-TeamChannel -GroupId $group.GroupId -DisplayName "Neuer Kanal" #>


#$UserCredential = Get-Credential
#$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection

#Import-PSSession $Session -DisableNameChecking

#Set-UnifiedGroup -Identity K3 -HiddenFromExchangeClientsEnabled:$False
#Remove-PSSession $Session


# https://www.tecklyfe.com/change-office-365-group-or-team-email-address/ #>

