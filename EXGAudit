#Be sure Powershell Remoting is enabled
#To enable on each server run winrm quickconfig

#Please save this file as exch.ps1 on one of your Exchange servers and save under C:\temp
#Launch Exchange Management Shell and run exch.ps1
#ZIP the output folder and email back to Concurrency

#Get Server Versions
$exchangeservers = Get-ExchangeServer
 
$report = @()
 
foreach ($srv in $exchangeservers){

$srv = Get-ExchangeServer $srv
 
$server = $srv.Name
if ($srv.AdminDisplayVersion -match "Version 14") {$ver = "V14"}
if ($srv.AdminDisplayVersion -match "Version 15") {$ver = "V15"}
 
    Write-Host "Checking $server"
 
    $installpath = $null
 
    try
    {
        $installpath = Invoke-Command –Computername $server -ScriptBlock {$env:ExchangeInstallPath} -ErrorAction STOP
    }
    catch
    {
        Write-Warning $_.Exception.Message
        $installpath = "Unable to connect to server"
    }

Write-host "Install Path is: " + $installpath
$Path = $installpath + "Bin\ExSetup.exe"
$fileversion = (Get-Command $Path).FileVersionInfo
Write-Host "File Version is: " + $FileVersion.fileversion

    $serverObj = New-Object PSObject
	$serverObj | Add-Member NoteProperty -Name "Server Name" -Value $server
	$serverObj | Add-Member NoteProperty -Name "Install Path" -Value $installpath
	$serverObj | Add-Member NoteProperty -Name "Server Role" -Value $srv.serverrole
	$serverObj | Add-Member NoteProperty -Name "Server Version" -Value $fileversion.fileversion
 
   
    $report += $serverObj  

Clear-variable srv
Clear-variable server
Clear-variable installpath
Clear-variable ver
}
 
$report | export-csv exchangeversions.csv -notype
start-Transcript serverurls.txt
#Get Server URLs
foreach ($i in $exchangeservers)
    {
        if ((Get-ExchangeServer $i -ErrorAction SilentlyContinue).IsClientAccessServer)
        {
            Write-Host "----------------------------------------"
            Write-Host " Querying $i"
            Write-Host "----------------------------------------`r`n"
            Write-Host "`r`n"

            $OA = Get-OutlookAnywhere -Server $i -AdPropertiesOnly | Select InternalHostName,ExternalHostName
            Write-Host "Outlook Anywhere"
            Write-Host " - Internal: $($OA.InternalHostName)"
            Write-Host " - External: $($OA.ExternalHostName)"
            Write-Host "`r`n"

            $OWA = Get-OWAVirtualDirectory -Server $i -AdPropertiesOnly | Select InternalURL,ExternalURL
            Write-Host "Outlook Web App"
            Write-Host " - Internal: $($OWA.InternalURL)"
            Write-Host " - External: $($OWA.ExternalURL)"
            Write-Host "`r`n"

            $ECP = Get-ECPVirtualDirectory -Server $i -AdPropertiesOnly | Select InternalURL,ExternalURL
            Write-Host "Exchange Control Panel"
            Write-Host " - Internal: $($ECP.InternalURL)"
            Write-Host " - External: $($ECP.ExternalURL)"
            Write-Host "`r`n"

            $OAB = Get-OABVirtualDirectory -Server $i -AdPropertiesOnly | Select InternalURL,ExternalURL
            Write-Host "Offline Address Book"
            Write-Host " - Internal: $($OAB.InternalURL)"
            Write-Host " - External: $($OAB.ExternalURL)"
            Write-Host "`r`n"

            $EWS = Get-WebServicesVirtualDirectory -Server $i -AdPropertiesOnly | Select InternalURL,ExternalURL
            Write-Host "Exchange Web Services"
            Write-Host " - Internal: $($EWS.InternalURL)"
            Write-Host " - External: $($EWS.ExternalURL)"
            Write-Host "`r`n"

            $MAPI = Get-MAPIVirtualDirectory -Server $i -AdPropertiesOnly | Select InternalURL,ExternalURL
            Write-Host "MAPI"
            Write-Host " - Internal: $($MAPI.InternalURL)"
            Write-Host " - External: $($MAPI.ExternalURL)"
            Write-Host "`r`n"

            $EAS = Get-ActiveSyncVirtualDirectory -Server $i -AdPropertiesOnly | Select InternalURL,ExternalURL
            Write-Host "ActiveSync"
            Write-Host " - Internal: $($EAS.InternalURL)"
            Write-Host " - External: $($EAS.ExternalURL)"
            Write-Host "`r`n"

	    #Depreciated	
            #$AutoD = Get-ClientAccessServer $i | Select AutoDiscoverServiceInternalUri
            #Write-Host "Autodiscover"
            #Write-Host " - Internal SCP: $($AutoD.AutoDiscoverServiceInternalUri)"
            #Write-Host "`r`n"

        }
        else
        {
            Write-Host -ForegroundColor Yellow "$i is not a Client Access server."
        }
    }
stop-transcript

Write-Host "Pull mailbox + db sizes (mailbox report)"
Get-Mailbox -Resultsize Unlimited | Get-MailboxStatistics | select-object DisplayName, alias, Database, {$_.TotalItemSize.Value.ToMB()}, ItemCount, {$_.TotalDeletedItemSize.Value.ToMB()}, DeletedItemCount, OrganizationalUnit, LastLogonTime | Export-CSV mbxDBsize.csv -notype

Write-Host "Pull quotas, policies (policy report)"
Get-Mailbox -Resultsize Unlimited | select-object Displayname, Alias, PrimarySMTPAddress, UserPrincipalName, RecipientTypeDetails, OrganizationalUnit, UseDatabaseQuotaDefaults, EmailAddressPolicyEnabled, *Litigation*, InPlaceHolds, RetentionPolicy, ManagedFolderMailboxPolicy, WhenMailboxCreated | Export-CSV mbxPOLICYlist.csv -notype

Write-Host "Pull all recipient types of Group (group)"
Get-Recipient -ResultSize Unlimited | ?{$_.RecipientType -like "*group*"} | select name, recipienttype, OrganizationalUnit, primarysmtpaddress | export-csv dllist.csv –notype

Write-Host "Pull Remote IP ranges for Receive connectors (remoteips)"
Get-ReceiveConnector | Select server,name -expandproperty RemoteIPRanges | Sort Name | Export-CSV remoteip.csv –notype

Write-Host "Pull Mailbox Full Access Permissions List"
Get-Mailbox -Resultsize Unlimited | Get-MailboxPermission | where {$_.user.tostring() -ne "NT AUTHORITY\SELF" -and $_.IsInherited -eq $false} | Select Identity,User,@{Name='Access Rights';Expression={[string]::join(', ', $_.AccessRights)}} | Export-Csv -NoTypeInformation mbxACLsource.csv

Write-Host "Pull Mailbox Delegate Permissions List"
Get-Mailbox -ResultSize unlimited | Get-CalendarProcessing | where { $_.ResourceDelegates -ne "" } | Select-Object identity,@{Name=’ResourceDelegates’;Expression={[string]::join(",", ($_.ResourceDelegates))}} | Export-csv -Path mbxResourceDelegates.csv 

Write-Host "Pull User Archive Mailboxes"
Get-Mailbox -Archive -Resultsize Unlimited | Get-MailboxStatistics | select-object DisplayName, alias, Database, {$_.TotalItemSize.Value.ToMB()}, ItemCount, {$_.TotalDeletedItemSize.Value.ToMB()}, DeletedItemCount, OrganizationalUnit, LastLogonTime | Export-CSV mbxARCsize.csv -notype

Write-Host "Pull Mailbox Databases sizes + white space"
Get-MailboxDatabase -Status | sort name | select name,@{Name='DB Size (Gb)';Expression={$_.DatabaseSize.ToGb()}},@{Name='Available New Mbx Space Gb)';Expression={$_.AvailableNewMailboxSpace.ToGb()}} | Export-csv mdb-size.csv –notype

Write-Host "Pull Database Core Data"
Get-MailboxDatabase | select Name, Server, JournalRecipient, OfflineAddressBook, PublicFolderDatabase, IsExcludedFromProvisioning, IsSuspendedFromProvisioning, EdbFilePath, LogFolderPath, CircularLoggingEnabled, rpcClientAccessServer, DeletedItesRetention | export-csv mbd-stat.csv -notype

Write-Host "Pull DAG Name"
Get-DatabaseAvailabilityGroup | select name,*network*,*activation* | Export-Csv DAGinfo.csv –notype

Write-Host "Pull DAG Network Settings"
Get-DatabaseAvailabilityGroupNetwork | select name, MapiAccessEnabled | Export-csv DAGnetwork.csv -notype

Write-Host "Pull activation preferences (DAG)"
Get-MailboxDatabase | select name,@{Name=' ActivationPreference ';Expression={[string]::join(', ', $_.ActivationPreference)}}, MasterServerOrAvailabilityGroup | export-csv mbd-activation.csv -notype

Write-Host "Pull ActiveSync Devices"
foreach ($user in $device) { Get-ActiveSyncDeviceStatistics -mailbox $user | Select-object identity, devicetype, Lastsuccesssync | Export-csv -encoding "unicode" -Path activesyncdevice.csv -NoType}

Write-Host "Pull Exchange Certificates
Get-ExchangeCertificate | Select @{Name='CertificateDomains';Expression={[string]::join(', ', $_.CertificateDomains)}}, HasPrivateKey, IsSelfSigned, Issuer, NotAfter, NotBefore, PublicKeySize, SerialNumber, Services, Status, Subject, Thumbprint | export-csv exg-cert.csv -notype

Write-Host "Pull Public Folder Database Info"
Get-publicfolderdatabase | Select Identity,Alias,@{Name='Servers';Expression={[string]::join(', ', $_.Servers)}},EdbFilePath,LogFolderPath | Export-Csv pub-info.csv -notype

Write-Host "Pull Public Folder Statistics"
Get-publicfolderstatistics -ResultSize Unlimited | select name,folderpath,lastaccesstime,lastmodificationtime, {$_.TotalItemSize.Value.ToKB()}, ItemCount, {$_.TotalDeletedItemSize.Value.ToKB()} | export-csv pub-stats.csv -notype

Write-Host "Pull Retention Policies"
Get-RetentionPolicy | select name,retentionpolicytaglinks | export-csv retention-policy.csv -notype
Get-RetentionPolicyTag | select "Name","Description","RetentionEnabled","RetentionAction","AgeLimitForRetention","MoveToDestinationFolder","TriggerForRetention" | export-csv retention-tag.csv -NoTypeInformation

Write-Host "Pull In-Place Hold Search Policies"
Get-mailboxsearch | select Name,@{Name='SourceMailboxes';Expression={[string]::join(', ', $_.SourceMailboxes)}},@{Name='Sources';Expression={[string]::join(', ', $_.Sources)}},@{Name='PublicFolderSources';Expression={[string]::join(', ', $_.PublicFolderSources)}},AllPublicFolderSources,@{Name='SearchStatistics';Expression={[string]::join(', ', $_.SearchStatistics)}},Version,@{Name='TargetMailbox';Expression={[string]::join(', ', $_.TargetMailbox)}},@{Name='Target';Expression={[string]::join(', ', $_.Target)}},@{Name='SearchQuery';Expression={[string]::join(', ', $_.SearchQuery)}},Language,@{Name='Senders';Expression={[string]::join(', ', $_.Senders)}},@{Name='Recipients';Expression={[string]::join(', ', $_.Recipients)}},StartDate,EndDate,@{Name='MessageTypes';Expression={[string]::join(', ', $_.MessageTypes)}},IncludeUnsearchableItems,EstimateOnly,ExcludeDuplicateMessages,Resume,IncludeKeywordStatistics,KeywordStatisticsDisabled,PreviewDisabled,@{Name='Information';Expression={[string]::join(', ', $_.Information)}},StatisticsStartIndex,TotalKeywords,LogLevel,@{Name='StatusMailRecipients';Expression={[string]::join(', ', $_.StatusMailRecipients)}},Status,LastRunBy,LastStartTime,LastEndTime,NumberMailboxesToSearch,PercentComplete,ResultNumber,ResultNumberEstimate,ResultSize,ResultSizeEstimate,ResultSizeCopied,ResultsLink,PreviewResultsLink,@{Name='Errors';Expression={[string]::join(', ', $_.Errors)}},InPlaceHoldEnabled,ItemHoldPeriod,InPlaceHoldIdentity,ManagedByOrganization,@{Name='FailedToHoldMailboxes';Expression={[string]::join(', ', $_.FailedToHoldMailboxes)}},@{Name='InPlaceHoldErrors';Expression={[string]::join(', ', $_.InPlaceHoldErrors)}},Description,LastModifiedTime,KeywordHits,IsValid,ObjectState | export-csv mbx-search.csv -notype

Write-Host "Data collection complete; please ZIP and return to consultant." -Foregroundcolor Yellow
