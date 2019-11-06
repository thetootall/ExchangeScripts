#Get Shared Mailbox Members into a CSV
#Chris Jorenby

#Change Window Title
$host.ui.RawUI.WindowTitle = "Get Shared Mailbox Members"
Write-Host "This Script will gather all shared mailboxes and their respected members" -ForegroundColor Yellow
Write-Host "I need your Office 365 Credentials" -ForegroundColor Yellow

#Get Login Credentials
$UserCredential = Get-Credential -Message "Enter your Office 365 Credentials"

#Make new session
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection

#Connect to PsSession
Import-PSSession $Session -WarningAction SilentlyContinue
Connect-MsolService -Credential $UserCredential

#Create A Table
$csvLine = New-Object PSObject
$csvTable = @()

#Get Shared Mailboxes
$SharedMailboxes = Get-Mailbox -ResultSize Unlimited | where {$_.RecipientTypeDetails -eq "RoomMailbox" -or $_.RecipientTypeDetails -eq "SharedMailbox"} | Select Name, SamAccountName, Office, UserPrincipalName, PrimarySmtpAddress, RecipientType, RecipientTypeDetails, MailTip, Identity, UseDatabaseRetentionDefaults, LitigationHoldEnabled, SingleItemRecoveryEnabled, LitigationHoldDuration, RetentionPolicy, ExchangeUserAccountControl, ResourceCapacity, ResourceCustom, ResourceType, RoomMailboxAccountEnabled, ThrottlingPolicy, RoleAssignmentPolicy, DefaultPublicFolderMailbox, EffectivePublicFolderMailbox, SharingPolicy, RemoteAccountPolicy, MailboxPlan, PersistedCapabilities, SKUAssigned, AuditEnabled, AuditLogAgeLimit, DefaultAuditSet, WhenMailboxCreated, UsageLocation, AccountDisabled, StsRefreshTokensValidFrom, EnforcedTimestamps, HasPicture, HasSpokenName, IsDirSynced, Alias, OrganizationalUnit, CustomAttribute1, CustomAttribute10, CustomAttribute11, CustomAttribute12, CustomAttribute13, CustomAttribute14, CustomAttribute15, CustomAttribute2, CustomAttribute3, CustomAttribute4, CustomAttribute5, CustomAttribute6, CustomAttribute7, CustomAttribute8, CustomAttribute9, ExtensionCustomAttribute1, ExtensionCustomAttribute2, ExtensionCustomAttribute3, ExtensionCustomAttribute4, ExtensionCustomAttribute5, DisplayName, GrantSendOnBehalfTo, HiddenFromAddressListsEnabled

$SharedMailboxes | export-csv ./sharedmbxList.csv -NoTypeInformation

$SharedMailboxes | foreach {
	$mailbox = Get-mailbox -Identity $_.Alias -ResultSize Unlimited
	$members = get-Mailboxpermission -Identity $mailbox.Alias | where {$_.User -like "*@*"}
	$memcount = $members.count
	$outname = "Gathering $mailbox Members"
	$outcount = "Showing $memcount Users"
	Write-Host $outname -ForegroundColor Cyan 
	Write-Host $outcount -ForegroundColor Yellow
	$outname | out-file ./sharedmbxLog.txt -Append
	$outcount | out-file ./sharedmbxLog.txt -Append
	$members | select Identity,User,AccessRights | export-csv ./sharedmbxExport.csv -Append -NoTypeInformation
}
