Write-host "Please save this file as exch.ps1 on one of your Exchange servers under a new output folder (IE C:\exgstats)" -BackgroundColor White -ForegroundColor Red
Write-host "Launch Exchange Management Shell and run exch.ps1" -BackgroundColor White -ForegroundColor Red
Write-host "ZIP the output folder and email back to Concurrency" -BackgroundColor White -ForegroundColor Red
Write-host " "
Write-host "Be sure Powershell Remoting is enabled on all Exchange servers." -BackgroundColor White -ForegroundColor Red
Write-host "If not, the first step, "Gathering Exchange Server build info", will fail. All other scripts should run." -BackgroundColor White -ForegroundColor Red
Write-host "To enable on each server run winrm quickconfig" -BackgroundColor White -ForegroundColor Red
Write-host " "
Pause

PAUSE

#Get Server Versions
Write-host "Gathering Exchange Server build info" -BackgroundColor Green -ForegroundColor Black

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
        $installpath = Invoke-Command “Computername $server -ScriptBlock {$env:ExchangeInstallPath} -ErrorAction STOP"
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
 
Write-host "Pull Exchange Server URLs" -BackgroundColor Green -ForegroundColor Black
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
            $OA = Get-OutlookAnywhere -Server $i -AdPropertiesOnly | Select InternalHostName,ExternalHostName,ClientAuthenticationMethod,IISAuthenticationMethods,ExternalClientAuthenticationMethod,InternalClientAuthenticationMethod
            Write-Host "Outlook Anywhere"
            Write-Host " - Internal: $($OA.InternalHostName)"
            Write-Host " - External: $($OA.ExternalHostName)"
	        Write-host " - Client Auth (Exchange 2010): $($OA.ClientAuthenticationMethod)"
	        Write-host " - Internal Client Auth (Exchange 2013): $($OA.InternalClientAuthenticationMethod)"
	        Write-host " - External Client Auth (Exchange 2013): $($OA.ExternalClientAuthenticationMethod)"
	        Write-host " - IIS Auth: $($OA.IISAuthenticationMethods)"
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

            $EWS = Get-WebServicesVirtualDirectory -Server $i -AdPropertiesOnly | Select InternalNLBBypassUrl,InternalURL,ExternalURL
            Write-Host "Exchange Web Services"
            Write-Host " - InternalNLBBypassUrl: $($EWS.InternalNLBBypassUrl)"
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

Write-host "Processing master user list, please wait..." -BackgroundColor White -ForegroundColor Black
Get-Mailbox -Resultsize Unlimited | select displayname,samaccountname,alias,primarysmtpadress | sort-object displayname | export-csv mbx-alluser.csv -notype
$alluser = import-csv mbx-alluser.csv

Write-Host "Pull mailbox + db sizes (mailbox report)" -BackgroundColor Green -ForegroundColor Black
#Added logic for percentage bar and read from master list plus output each user at a time
#Get-Mailbox -Resultsize Unlimited | Get-MailboxStatistics | select-object DisplayName, alias, Database, {$_.TotalItemSize.Value.ToMB()}, ItemCount, {$_.TotalDeletedItemSize.Value.ToMB()}, DeletedItemCount, OrganizationalUnit, LastLogonTime | Export-CSV mbxDBsize.csv -notype

$numstat = 0
foreach ($i in $alluser){
$myuser = $i.samaccountname
$mydisp = $i.displayname
Write-host "Reading info for $mydisp"

#Output the details and show progress
Get-MailboxStatistics $myuser | select-object DisplayName, Database, {$_.TotalItemSize.Value.ToMB()}, ItemCount, {$_.TotalDeletedItemSize.Value.ToMB()}, DeletedItemCount, OrganizationalUnit, LastLogonTime | Export-CSV mbxDBsize.csv -append
Write-Progress -Activity "Outputting User Statistics" -Status "Progress:" -PercentComplete ($numstat/$alluser.count*100)
$numstat = $numstat+1
}

Write-Host "Pull Mailbox Full Access Permissions List" -BackgroundColor Green -ForegroundColor Black
#Added logic for percentage bar and read from master list plus output each user at a time
#Get-Mailbox -ResultSize unlimited | Get-CalendarProcessing | where { $_.ResourceDelegates -ne "" } | Select-Object identity,@{Name="ResourceDelegates";Expression={[string]::join(",", ($_.ResourceDelegates))}} | Export-csv -Path mbxResourceDelegates.csv 
$numperm = 0
foreach ($in in $alluser){
$mypermuser = $in.samaccountname
$mypermdisp = $in.displayname
Write-host "Reading info for $mypermdisp"

#Output the details and show progress
Get-MailboxPermission $mypermuser | where {$_.user.tostring() -ne "NT AUTHORITY\SELF" -and $_.IsInherited -eq $false} | Select Identity,User,@{Name='Access Rights';Expression={[string]::join(', ', $_.AccessRights)}} | Export-Csv -NoTypeInformation mbxACLsource.csv -append
Write-Progress -Activity "Outputting Mailox Full Access Permissions" -Status "Progress:" -PercentComplete ($numperm/$alluser.count*100)
$numperm = $numperm+1
}

Write-Host "Pull Mailbox Delegate Permissions List" -BackgroundColor Green -ForegroundColor Black
#Added logic for percentage bar and read from master list plus output each user at a time
#Get-Mailbox -ResultSize unlimited | Get-CalendarProcessing | where { $_.ResourceDelegates -ne "" } | Select-Object identity,@{Name="ResourceDelegates";Expression={[string]::join(",", ($_.ResourceDelegates))}} | Export-csv -Path mbxResourceDelegates.csv 
$numdell = 0
foreach ($ind in $alluser){
$mydeluser = $ind.samaccountname
$mydeldisp = $ind.displayname
Write-host "Reading info for $mydeldisp"

#Output the details and show progress
Get-CalendarProcessing $mydeluser | where { $_.ResourceDelegates -ne "" } | Select-Object identity,@{Name="ResourceDelegates";Expression={[string]::join(",", ($_.ResourceDelegates))}} | Export-csv -Path mbxResourceDelegates.csv -append
Write-Progress -Activity "Outputting Mailbox Delegate Permissions" -Status "Progress:" -PercentComplete ($numdell/$alluser.count*100)
$numperm = $numperm+1
}

Write-Host "Pull User Archive Mailboxes" -BackgroundColor Green -ForegroundColor Black
$alluserarc = Get-Mailbox -Archive -Resultsize Unlimited | select displayname,samaccountname | sort-order displayname
$numarc = 0
foreach ($inc in $alluser){
$myarcuser = $inc.samaccountname
$myarcdisp = $inc.displayname
Write-host "Reading info for $myarcdisp"

#Output the details and show progress
Get-MailboxStatistics $myarcuser | select-object DisplayName, alias, Database, {$_.TotalItemSize.Value.ToMB()}, ItemCount, {$_.TotalDeletedItemSize.Value.ToMB()}, DeletedItemCount, OrganizationalUnit, LastLogonTime | Export-CSV mbxARCsize.csv -append
Write-Progress -Activity "Outputting User Archive Mailboxes" -Status "Progress:" -PercentComplete ($numarc/$alluser.count*100)
$numarc = $numarc+1
}


Write-Host "Pull quotas, policies (policy report)" -BackgroundColor Green -ForegroundColor Black
#Added logic for percentage bar and read from master list plus output each user at a time
$numpol = 0
foreach ($inp in $alluser){
$mypoluser = $inp.samaccountname
$mypoldisp = $inp.displayname
Write-host "Reading info for $mypoldisp"

#Output the details and show progress
Get-Mailbox $mypoluser | select-object Displayname, Alias, PrimarySMTPAddress, UserPrincipalName, RecipientTypeDetails, OrganizationalUnit, UseDatabaseQuotaDefaults, EmailAddressPolicyEnabled, *Litigation*, InPlaceHolds, RetentionPolicy, ManagedFolderMailboxPolicy, WhenMailboxCreated | Export-CSV mbxPOLICYlist.csv -append
Write-Progress -Activity "Outputting User Mailbox Policies" -Status "Progress:" -PercentComplete ($numpol/$alluser.count*100)
$numpol = $numpol+1
}

Write-Host "Pull all recipient types of Group (group)" -BackgroundColor Green -ForegroundColor Black
Get-Recipient -ResultSize Unlimited | ?{$_.RecipientType -like "*group*"} | select name, recipienttype, OrganizationalUnit, primarysmtpaddress | export-csv dllist.csv -notype

Write-Host "Pull Remote IP ranges for Receive connectors (remoteips)" -BackgroundColor Green -ForegroundColor Black
Get-ReceiveConnector | Select server,name -expandproperty RemoteIPRanges | Sort Name | Export-CSV remoteip.csv -notype

Write-host "Pull Send Connector information" -BackgroundColor Green -ForegroundColor Black
Get-SendConnector | list | out-file sendconnector.txt

Write-Host "Pull Mailbox Databases sizes + white space" -BackgroundColor Green -ForegroundColor Black
Get-MailboxDatabase -Status | sort name | select name,@{Name='DB Size (Gb)';Expression={$_.DatabaseSize.ToGb()}},@{Name='Available New Mbx Space Gb)';Expression={$_.AvailableNewMailboxSpace.ToGb()}} | Export-csv mdb-size.csv -notype

Write-Host "Pull Database Core Data" -BackgroundColor Green -ForegroundColor Black
Get-MailboxDatabase | select Name, Server, JournalRecipient, OfflineAddressBook, PublicFolderDatabase, IsExcludedFromProvisioning, IsSuspendedFromProvisioning, EdbFilePath, LogFolderPath, CircularLoggingEnabled, rpcClientAccessServer, DeletedItemRetention, SendReceiveQuota, SendQuota, WarningQuota | export-csv mbd-stat.csv -notype

Write-Host "Pull DAG Name" -BackgroundColor Green -ForegroundColor Black
Get-DatabaseAvailabilityGroup | select name,*network*,*activation* | Export-Csv DAGinfo.csv -notype

Write-Host "Pull DAG Network Settings" -BackgroundColor Green -ForegroundColor Black
Get-DatabaseAvailabilityGroupNetwork | select name, MapiAccessEnabled | Export-csv DAGnetwork.csv -notype

Write-Host "Pull activation preferences (DAG)" -BackgroundColor Green -ForegroundColor Black
Get-MailboxDatabase | select name,@{Name=' ActivationPreference ';Expression={[string]::join(', ', $_.ActivationPreference)}}, MasterServerOrAvailabilityGroup | export-csv mbd-activation.csv -notype

Write-Host "Pull ActiveSync Devices" -BackgroundColor Green -ForegroundColor Black
foreach ($user in $device) { Get-ActiveSyncDeviceStatistics -mailbox $user | Select-object identity, devicetype, Lastsuccesssync | Export-csv -encoding "unicode" -Path activesyncdevice.csv -NoType}

Write-Host "Pull Exchange Certificates" -BackgroundColor Green -ForegroundColor Black
Get-ExchangeCertificate | Select @{Name='CertificateDomains';Expression={[string]::join(', ', $_.CertificateDomains)}}, HasPrivateKey, IsSelfSigned, Issuer, NotAfter, NotBefore, PublicKeySize, SerialNumber, Services, Status, Subject, Thumbprint | export-csv exg-cert.csv -notype

Write-Host "Pull Public Folder Database Info" -BackgroundColor Green -ForegroundColor Black
Get-publicfolderdatabase | Select Identity,Alias,@{Name='Servers';Expression={[string]::join(', ', $_.Servers)}},EdbFilePath,LogFolderPath | Export-Csv pub-info.csv -notype

Write-Host "Pull Public Folder Statistics" -BackgroundColor Green -ForegroundColor Black
Get-publicfolderstatistics -ResultSize Unlimited | select name,folderpath,lastaccesstime,lastmodificationtime, {$_.TotalItemSize.Value.ToKB()}, ItemCount, {$_.TotalDeletedItemSize.Value.ToKB()} | export-csv pub-stats.csv -notype

Write-Host "Pull Retention Policies" -BackgroundColor Green -ForegroundColor Black
Get-RetentionPolicy | select name,retentionpolicytaglinks | export-csv retention-policy.csv -notype
Get-RetentionPolicyTag | select "Name","Description","RetentionEnabled","RetentionAction","AgeLimitForRetention","MoveToDestinationFolder","TriggerForRetention" | export-csv retention-tag.csv -NoTypeInformation

Write-Host "Pull In-Place Hold Search Policies" -BackgroundColor Green -ForegroundColor Black
Get-mailboxsearch | select Name,@{Name='SourceMailboxes';Expression={[string]::join(', ', $_.SourceMailboxes)}},@{Name='Sources';Expression={[string]::join(', ', $_.Sources)}},@{Name='PublicFolderSources';Expression={[string]::join(', ', $_.PublicFolderSources)}},AllPublicFolderSources,@{Name='SearchStatistics';Expression={[string]::join(', ', $_.SearchStatistics)}},Version,@{Name='TargetMailbox';Expression={[string]::join(', ', $_.TargetMailbox)}},@{Name='Target';Expression={[string]::join(', ', $_.Target)}},@{Name='SearchQuery';Expression={[string]::join(', ', $_.SearchQuery)}},Language,@{Name='Senders';Expression={[string]::join(', ', $_.Senders)}},@{Name='Recipients';Expression={[string]::join(', ', $_.Recipients)}},StartDate,EndDate,@{Name='MessageTypes';Expression={[string]::join(', ', $_.MessageTypes)}},IncludeUnsearchableItems,EstimateOnly,ExcludeDuplicateMessages,Resume,IncludeKeywordStatistics,KeywordStatisticsDisabled,PreviewDisabled,@{Name='Information';Expression={[string]::join(', ', $_.Information)}},StatisticsStartIndex,TotalKeywords,LogLevel,@{Name='StatusMailRecipients';Expression={[string]::join(', ', $_.StatusMailRecipients)}},Status,LastRunBy,LastStartTime,LastEndTime,NumberMailboxesToSearch,PercentComplete,ResultNumber,ResultNumberEstimate,ResultSize,ResultSizeEstimate,ResultSizeCopied,ResultsLink,PreviewResultsLink,@{Name='Errors';Expression={[string]::join(', ', $_.Errors)}},InPlaceHoldEnabled,ItemHoldPeriod,InPlaceHoldIdentity,ManagedByOrganization,@{Name='FailedToHoldMailboxes';Expression={[string]::join(', ', $_.FailedToHoldMailboxes)}},@{Name='InPlaceHoldErrors';Expression={[string]::join(', ', $_.InPlaceHoldErrors)}},Description,LastModifiedTime,KeywordHits,IsValid,ObjectState | export-csv mbx-search.csv -notype

Write-Host "Data collection complete; please ZIP and return to consultant." -Foregroundcolor Yellow
