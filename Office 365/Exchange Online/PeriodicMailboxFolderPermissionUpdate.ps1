#ToDo:
#	- Get only mailboxes in a particular group (eg, Office, Contractors, Rooms)
#	- Don't change permissions if it is a level higher than intended setting (eg, owner?)

Param(
	$AzureAutomationCredential = "Exchange Online admin",
	$GroupsToApplyTo = @("Office", "Contractors", "Rooms"),
	$GroupsToAssignPermissionTo = @("Office","Contractors"),
	$AccessRights = "Reviewer",
	$AllowableAccessRights = @("Owner","Publishing Editor","Editor","Publishing Author","Author","Nonediting Author"),
	$Folder = "Calendar"
);

Write-Output "Getting credentials $AzureAutomationCredential..."
$cred = Get-AutomationPSCredential -Name $AzureAutomationCredential;

Write-Output "Targetted AccessRights: $AccessRights"
$AllowableAccessRights += $AccessRights;
Write-Output ("AllowableAccessRights: " + ($AllowableAccessRights -join ", ") );
Write-Output "Targetted Folder: $Folder"

Write-Output "Connecting to Exchange Online PSSession..."
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $cred -Authentication Basic -AllowRedirection;
Import-PSSession $Session;

#$AllMailboxes = Get-Mailbox -ResultSize Unlimited;
#$AllMailboxes = Get-Mailbox -ResultSize Unlimited -Filter {(MemberOfGroup -eq "")};
#$AllMailboxes = Get-Mailbox -Identity "Mark.Grose"

$AllGroupMembers = @();
$AllMailboxesAliases = @();

Write-Output ("GroupsToApplyTo they have the permissions added: " + ($GroupsToApplyTo -join ", "))
$GroupsToApplyTo | % {
	Write-Verbose "`t$_";
	$AllGroupMembers += Get-DistributionGroupMember $_ -ResultSize Unlimited;
	Write-Verbose ("`t`t" + (($AllGroupMembers | Select-Object -ExpandProperty Alias | Sort-Object) -join ", "));
}

Write-Output "Collecting all unique mailbox aliases..." 
$AllMailboxesAliases = $AllGroupMembers | Select-Object -Unique alias -ExpandProperty alias | Sort-Object | ? { $_ -ne $null -and $_ -ne "" }

Write-Output "Processing mailboxes now..."
ForEach ( $group in $GroupsToAssignPermissionTo ) {
	Foreach ($MailboxAlias in $AllMailboxesAliases) {
		Write-Verbose "Getting existing $folder permission for $MailboxAlias..."

		$ExistingPermission = $null;
		$ExistingPermission = Get-MailboxFolderPermission -Identity ($MailboxAlias+':\'+$Folder) -User $Group -ErrorAction SilentlyContinue;
		
		if ( $ExistingPermission -eq $null ) {
			Write-Output ("Identity {0} does not have {1} access rights on their {3} for group '{2}'...adding..." -f $MailboxAlias, $AccessRights, $group, $Folder);
			Write-Output ("Add-MailboxFolderPermission -Identity $($MailboxAlias+':\'+$Folder) -User $group -AccessRights $AccessRights;");
			Add-MailboxFolderPermission -Identity ($MailboxAlias+':\'+$Folder) -User $group -AccessRights $AccessRights;
		} elseif ( ($ExistingPermission.AccessRights | ? { $AllowableAccessRights.Contains($_) }) -eq $null ) {
			#ToDo: Deal properly with multiple AccessRights, perhaps just use Add-?
			Write-Warning ("Identity {0} does not have allowable access rights on {4} folder for group '{2}', instead had {3}...reconfiguring..." -f $MailboxAlias, $group, $ExistingPermission.AccessRights, $Folder);
			Write-Output ("Set-MailboxFolderPermission -Identity $($MailboxAlias+':\'+$Folder) -User $group -AccessRights $AccessRights;");
			Add-MailboxFolderPermission -Identity ($MailboxAlias+':\'+$Folder) -User $group -AccessRights $AccessRights;
		} else {
			Write-Verbose ("No change required for group {2} on identity {0}'s folder {1} as its set as {3}" -f $MailboxAlias, $Folder, $group, $ExistingPermission.AccessRights);
		}
	}
}