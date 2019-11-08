#Connect to all O365 modules
#https://www.powershellgallery.com/packages/Posh365/0.3.4/Content/Public%5CConnect%5CConnect-Cloud.ps1
#https://www.thelazyadministrator.com/2019/02/05/powershell-function-to-connect-to-all-office-365-services-with-support-for-mfa/

#ensure that WiNRM is Enabled
winrm set winrm/config/client/auth @{Basic="true"}
winrm quickconfig

#Import the module, requires that you are administrator and are able to run the script
#Thanks to Vlad for the extra cmds to account for MFA
Install-Module -Name CreateExoPsSession -Scope CurrentUser -Force
Import-Module $((Get-ChildItem -Path $($env:LOCALAPPDATA+"\Apps\2.0\") -Filter CreateExoPSSession.ps1 -Recurse ).FullName | Select-Object -Last 1)
#connect specifying username, if you already have authenticated to another module, you actually do not have to authenticate
$myupn = Read-Host "Please enter your administrator UPN"
Connect-EXOPSSession -UserPrincipalName $myupn
#This will make sure when you need to reauthenticate after 1 hour that it uses existing token and you don't have to write password and stuff
$global:UserPrincipalName=$myupn

#Get Shared Mailbox Members into a CSV
#Chris Jorenby

#Change Window Title
$host.ui.RawUI.WindowTitle = "Get Shared Mailbox Members"
Write-Host "This Script will gather all shared mailboxes and their respected members" -ForegroundColor Yellow
Write-Host "I need your Office 365 Credentials" -ForegroundColor Yellow

#Create A Table
$csvLine = New-Object PSObject
$csvTable = @()

#Get Shared Mailboxes
$SharedMailboxes = Get-Mailbox -ResultSize Unlimited | where {$_.RecipientTypeDetails -eq "RoomMailbox" -or $_.RecipientTypeDetails -eq "SharedMailbox"} | Select Name, SamAccountName, Office, UserPrincipalName, PrimarySmtpAddress, RecipientType, RecipientTypeDetails, MailTip, Identity, UseDatabaseRetentionDefaults, LitigationHoldEnabled, SingleItemRecoveryEnabled, LitigationHoldDuration, RetentionPolicy, ExchangeUserAccountControl, ResourceCapacity, ResourceCustom, ResourceType, RoomMailboxAccountEnabled, ThrottlingPolicy, RoleAssignmentPolicy, DefaultPublicFolderMailbox, EffectivePublicFolderMailbox, SharingPolicy, RemoteAccountPolicy, MailboxPlan, PersistedCapabilities, SKUAssigned, AuditEnabled, AuditLogAgeLimit, DefaultAuditSet, WhenMailboxCreated, UsageLocation, AccountDisabled, StsRefreshTokensValidFrom, EnforcedTimestamps, HasPicture, HasSpokenName, IsDirSynced, Alias, OrganizationalUnit, CustomAttribute1, CustomAttribute10, CustomAttribute11, CustomAttribute12, CustomAttribute13, CustomAttribute14, CustomAttribute15, CustomAttribute2, CustomAttribute3, CustomAttribute4, CustomAttribute5, CustomAttribute6, CustomAttribute7, CustomAttribute8, CustomAttribute9, ExtensionCustomAttribute1, ExtensionCustomAttribute2, ExtensionCustomAttribute3, ExtensionCustomAttribute4, ExtensionCustomAttribute5, DisplayName, GrantSendOnBehalfTo, HiddenFromAddressListsEnabled
#Exports Shared Mailbox attributes
$SharedMailboxes | export-csv c:\temp\sharedmbxList.csv -NoTypeInformation

#Creates the Mailbox Size file and header
$header = "DisplayName|MBType|LastLogon|LastLogoff|TotalItems|TotalSizeInMB|DeletedItems|DeletedItemsInMB"
Write-host $header
$header | Out-file c:\temp\sharedmbxSize.txt -Append

$SharedMailboxes | foreach {
	$mailbox = Get-mailbox -Identity $_.Alias -ResultSize Unlimited
	$members = get-Mailboxpermission -Identity $mailbox.Alias | where {$_.User -like "*@*"}
	$membersize = 
	$memcount = $members.count
	$outname = "Gathering $mailbox Members"
	$outcount = "Showing $memcount Users"
	Write-Host $outname -ForegroundColor Cyan 
	Write-Host $outcount -ForegroundColor Yellow
	$outname | out-file c:\temp\sharedmbxLog.txt -Append
	$outcount | out-file c:\temp\sharedmbxLog.txt -Append
	$members | select Identity,User,AccessRights | export-csv c:\temp\sharedmbxExport.csv -Append -NoTypeInformation

$mymailbox = $mailbox.DistinguishedName
$mymailboxstat = $mymailbox | Get-MailboxStatistics | select userprincipalname,displayname,MailboxTypeDetail,lastlogontime,lastlogofftime,ItemCount,TotalItemSize,TotalDeletedItemSize,DeletedItemCount

$mymaildisp = $mymailboxstat.displayname
$mymailtype = $mymailboxstat.MailboxTypeDetail
$mymaillogon = $mymailboxstat.lastlogontime
$mymaillogoff = $mymailboxstat.lastlogofftime
$mymailsizeitem = $mymailboxstat.ItemCount
$mymailsizedeleted = $mymailboxstat.DeletedItemCount

$mymailitem = [math]::Round(($mymailboxstat.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB),2)
$mymaildeleted = [math]::Round(($mymailboxstat.TotalDeletedItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB),2)

$output = "$mymaildisp|$mymailtype|$mymaillogon|$mymaillogoff|$mymailitem|$mymailsizeitem|$mymaildeleted|$mymailsizedeleted"

Write-host $Output
$output | Out-file c:\temp\sharedmbxSize.txt -Append

}
