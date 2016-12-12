#ToDo:
#	- Get only mailboxes in a particular group (eg, Office, Contractors, Rooms)
#	- Don't change permissions if it is a level higher than intended setting (eg, owner?)

Param(
	$AzureAutomationCredential = "",
	$GroupsToEnsure = @("Office","Contractors"),
	$AccessRights = "Reviewer",
	$AllowableAccessRights = @("Owner","Contributor"),
	#$IdentityGroups = "Office", "Contractors", "Rooms", - not yet in use
	$Folder = "Calendar"
);

$cred = Get-AutomationPSCredential -Name $AzureAutomationCredential;

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $cred -Authentication Basic -AllowRedirection;
Import-PSSession $Session;

$AllowableAccessRights+=$AccessRights;

#$AllMailboxes = Get-Mailbox -ResultSize Unlimited;
$AllMailboxes = Get-Mailbox -Identity "Mark.Grose"

ForEach ( $group in $GroupsToHaveReviewer ) {
	Foreach ($Mailbox in $AllMailboxes) {
		$ExistingPermission = $null;
		$ExistingPermission = Get-MailboxFolderPermission -Identity ($Mailbox.Alias+':\'+$Folder) -User $Group -ErrorAction SilentlyContinue;
		
		if ( $ExistingPermission -eq $null ) {
			Write-Output ("Identity {0} does not have {1} access rights on their {3} for group '{2}'...adding..." -f $mailbox.alias, $AccessRights, $Folder);
			Write-Output ("`tAdd-MailboxFolderPermission -Identity $($Mailbox.Alias+':\'+$Folder) -User $group -AccessRights $AccessRights;");
			Add-MailboxFolderPermission -Identity ($Mailbox.Alias+':\'+$Folder) -User $group -AccessRights $AccessRights;
		} elseif ( $ExistingPermission | ? { $AllowableAccessRights.Contains($_.AccessRights) } -eq $null ) {
			#ToDo: Check .contains actually works here for this
			#ToDo: Check if there are multiple ExistingPermissions (if possible?), and deal with accordingly
			Write-Warning ("Identity {0} does not have {1} access rights on {4} folder for group '{2}', instead had {3}...reconfiguring..." -f $mailbox.alias, $AllowableAccessRights -join ", ", $group, $ExistingPermission.AccessRights, $Folder);
			Write-Output ("`tSet-MailboxFolderPermission -Identity $($Mailbox.Alias+':\'+$Folder) -User $group -AccessRights $AccessRights;");
			Set-MailboxFolderPermission -Identity ($Mailbox.Alias+':\'+$Folder) -User $group -AccessRights $AccessRights;
		} else {
			Write-Verbose ("No change required for group {2} on identity {0}'s folder {1}" -f $Mailbox.Alias, $Folder, $group)
		}
	}
}