<#  
.SYNOPSIS
	Ensures AccessRights (or AllowableAccessRights) permissions exist on GroupIdentityFolderOwner identities Exchange folder (eg, Calendar) for groups specified in PermissionBenefactorGroups 
.DESCRIPTION  
	Ensures AccessRights (or AllowableAccessRights) permissions exist on GroupIdentityFolderOwner identities Exchange folder (eg, Calendar) for groups specified in PermissionBenefactorGroups
.NOTES  
    Author:		ashley.geek.nz
	Version:	v1.1 2016-12-15 1830 (UTC+13)
	Intended to be run in Azure Automation periodically to enforce the settings, and as such uses Get-AutomationPSCredential
.PARAMETER AzureAutomationCredential
	The name of the Credential stored in the Azure Automation assets.
.PARAMETER GroupIdentityFolderOwner
	Distribution Group(s) of user/identity/mailboxes who own the Folder intended to have the permissions/AccessRights assigned
	Provided as an array of strings: ("Office", "Contractors")
.PARAMETER PermissionBenefactorGroups
	Security Group(s) which benefit from the permission/AccessRight being assigned to GroupIdentityFolderOwner's Folder.
	Provided as an array of strings: ("Office", "Contractors")
.PARAMETER AccessRights
	The named AccessRights Exchange permission to be applied to the Folder for each PermissionBenefactorGroup(s)
.PARAMETER AllowableAccessRights
	The permissions/AccessRights which are further/a superset to what is being enforced, and therefore are acceptable to use instead.
	Provided as an array of strings; eg, for Reviewer: @("Owner","Publishing Editor","Editor","Publishing Author","Author","Nonediting Author")
.PARAMETER Folder
	The Exchange folder in GroupIdentityFolderOwner's mailbox that should have the AccessRights applied to (eg, Calendar)
.LINK
	https://github.com/cc-ashley/automation/
#> 

Param(
	#[Parameter(Mandatory=$true)]
	[string]$AzureAutomationCredential = "Exchange Online admin",

	#[Parameter(Mandatory=$true)]
	[array]$GroupIdentityFolderOwner = ("Office", "Contractors"),

	#[Parameter(Mandatory=$true)]
	[array]$PermissionBenefactorGroups = ("Office","Contractors"),

	#[Parameter(Mandatory=$true)]
	[string]$AccessRights = "Reviewer",

	#[Parameter(Mandatory=$true)]
	[array]$AllowableAccessRights = ("Owner","Publishing Editor","Editor","Publishing Author","Author","Nonediting Author"),

	#[Parameter(Mandatory=$true)]
	[string]$Folder = "Calendar"
);

Write-Output "Getting credentials $AzureAutomationCredential..."
$cred = Get-AutomationPSCredential -Name $AzureAutomationCredential;

Write-Output "Targetted AccessRights: $AccessRights"
$AllowableAccessRights += $AccessRights;
Write-Output ("AllowableAccessRights: " + ($AllowableAccessRights -join ", ") );
Write-Output "Targetted Folder: $Folder"
Write-Output "";

Write-Output "Connecting to Exchange Online PSSession...";
$ConnUri = [System.Uri]("https://outlook.office365.com/powershell-liveid/");
$ExchOnlineSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ConnUri -Credential $cred -Authentication "Basic" -AllowRedirection;
Write-Output "Importing PSSession/Module...";
Import-Module (Import-PSSession -Session $ExchOnlineSession -DisableNameChecking -AllowClobber -WarningAction SilentlyContinue) -Global -WarningAction SilentlyContinue;

$AllGroupMembers = @();
$AllMailboxesAliases = @();

Write-Output ("GroupsToApply(ThePermissions)To: " + ($GroupIdentityFolderOwner -join ", "))
$GroupIdentityFolderOwner | % {
	Write-Verbose "`t$_";
	$AllGroupMembers += Get-DistributionGroupMember $_ -ResultSize Unlimited;
	Write-Verbose ("`t`t" + (($AllGroupMembers | Select-Object -ExpandProperty Alias | Sort-Object) -join ", "));
}
Write-Output ("PermissionBenefactorGroups: " + ($PermissionBenefactorGroups -join ", "))

Write-Output "Collecting all unique mailbox aliases..." 
$AllMailboxesAliases = $AllGroupMembers | Select-Object -Unique alias -ExpandProperty alias | Sort-Object | ? { $_ -ne $null -and $_ -ne "" }

Write-Output "Processing PermissionBenefactorGroups now..."
ForEach ( $group in $PermissionBenefactorGroups ) {
	Write-Output "`t$group...iterating mailboxes..."
	Foreach ($MailboxAlias in $AllMailboxesAliases) {
		Write-Verbose "`t`tGetting existing $folder permission for $MailboxAlias..."

		$ExistingPermission = $null;
		$ExistingPermission = Get-MailboxFolderPermission -Identity ($MailboxAlias+':\'+$Folder) -User $Group -ErrorAction SilentlyContinue;
		
		if ( $ExistingPermission -eq $null ) {
			$Output = ("Identity {0} does not have {1} access rights on their {3} for group '{2}'...adding..." -f $MailboxAlias, $AccessRights, $group, $Folder)
			Write-Warning $Output;
			Write-Output $Output;
			Write-Output ("Add-MailboxFolderPermission -Identity $($MailboxAlias+':\'+$Folder) -User $group -AccessRights $AccessRights;");
			Add-MailboxFolderPermission -Identity ($MailboxAlias+':\'+$Folder) -User $group -AccessRights $AccessRights;
			Write-Output "";
		} elseif ( ($ExistingPermission.AccessRights | ? { $AllowableAccessRights.Contains($_) }) -eq $null ) {
			$Output = ("Identity {0} does not have allowable access rights on {1} folder for group '{2}', instead had {3}...reconfiguring..." -f $MailboxAlias, $Folder, $group, $ExistingPermission.AccessRights)
			Write-Warning $Output;
			Write-Output $Output;
			Write-Output ("Set-MailboxFolderPermission -Identity $($MailboxAlias+':\'+$Folder) -User $group -AccessRights $AccessRights;");
			Set-MailboxFolderPermission -Identity ($MailboxAlias+':\'+$Folder) -User $group -AccessRights $AccessRights;
			Write-Output "";
		} else {
			Write-Verbose ("No change required for group {2} on identity {0}'s folder {1} as its set as {3}" -f $MailboxAlias, $Folder, $group, $ExistingPermission.AccessRights);
		}
	}
}

Write-Output "Completed"