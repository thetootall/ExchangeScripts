<#
.SYNOPSIS
Use this script to generate a migration schedule for mailboxes being migrated from on-premises environments to O365. The script will first collect all permissions, identify associated delegates, and generate a schedule in the format expected by the Microsoft FastTrack Center (FTC) team.
You can use the generated schedule to manipulate the mailboxes that you would like to migrate and share with the FTC team to submit for migration.  

.DESCRIPTION
The sample scripts are not supported under any Microsoft standard support 
program or service. The sample scripts are provided AS IS without warranty  
of any kind. Microsoft further disclaims all implied warranties including,  
without limitation, any implied warranties of merchantability or of fitness for 
a particular purpose. The entire risk arising out of the use or performance of  
the sample scripts and documentation remains with you. In no event shall 
Microsoft, its authors, or anyone else involved in the creation, production, or 
delivery of the scripts be liable for any damages whatsoever (including, 
without limitation, damages for loss of business profits, business interruption, 
loss of business information, or other pecuniary loss) arising out of the use 
of or inability to use the sample scripts or documentation, even if Microsoft 
has been advised of the possibility of such damages.

=========================================
Published date: 12/08/2016
+01252017: Corrected text in help from UserImportCSV to UseImportCSV 
+12152016: Added check for Get-User when creating the migration schedule
+12082016: Change order of Remove Service Accounts prompt
+07282016: Added FullAccess permissions. Disabled by default. 
+06272016: Update method to get calendar variable for different language
+06162016: Replaced preloaded variable for calendar with method to avoid issue with different language. Also updated Send As permissions to not include inherited permissions. 
+03212016: Make calendar folder permissions disabled by default. To enable set the following setting to $true -> $gathercalendar = $TRUE
+03012016: Included ListActiveSyncDevices and AddActiveSyncDevices functionality. Included guidance in the read me file. Updated to .docx 
+03012016: Removed FullAccess permissions, inserted Calendar-Folder permissions. Get Calendar Folder DISABLED BY DEFAULT.  
+01272016: Added condition to exclude exchange administration default groups when enumerating group membership.
+01262016: Added condition to catch exception for ManagementObjectNotFoundException. Get-Recipient logs error when object doesn't have email, leading to confusion. This is expected.  
+01222016: Updated scriptpath logic and $myinvocation to avoid intermittent split-path issues. 
+01152016: Added logging to spot potential perf issues easier
+01132016: Comment out ListActiveSyncDevices and AddActiveSyncDevices, these are in development
+12172015: Add PageSize
+12032015: Added check for Invocation. If it fails save files to c: drive. 
+11162015: Renamed OBC to FTC
+08112015: Group enumeration for both FullAccess and Send As Permissions
+08112015: Ability to list out ActiveSync Devices and to add to migrated mailboxes

Authors: 
Alejandro Lopez - alejanl@microsoft.com
Sam Portelli - Sam.Portelli@microsoft.com
=========================================

.PARAMETER UseImportCSV
Use this parameter to specify a list of users to collect permissions for, rather than all mailboxes.
Make sure that the CSV file provided has a header titled "Email"

.PARAMETER RemoveServiceAccts
Use this parameter to specify a list of service accounts (or even public shared mailboxes that all users have permissions to) that you would like to exclude from permissions
Make sure that the CSV file provided has a header titled "Email"

.PARAMETER ListActiveSyncDevices  
Use this parameter to include active sync devices allowed for the mailboxes in the on-premises environment. This will add a separate column to the migration schedule and list those devices in pipe-delimited format. 

.PARAMETER AddActiveSyncDevices 
Use this parameter to update migrated cloud mailboxes (using a schedule file of migrated mailboxes - must be schedule that includes the active sync devices, can be generated using the -ListActiveSyncDevices parameter) and add the active sync devices to the cloud mailboxes

.EXAMPLE
#Create Migration Schedule for ALL users - by default
.\ftcmigrationplanning.ps1 

.EXAMPLE
#Create Migration Schedule for ALL users EXCEPT service accounts
.\ftcmigrationplanning.ps1 -RemoveServiceAccts

.EXAMPLE
#Create Migration Schedule for list of users
.\ftcmigrationplanning.ps1 -UseImportCSV 

.EXAMPLE
#Create Migration Schedule for list of users except service accounts
.\ftcmigrationplanning.ps1 -UseImportCSV -RemoveServiceAccts

#>
    
param(
	[switch]$UseImportCSV,
	[switch]$RemoveServiceAccts,
    [switch]$ListActiveSyncDevices, 
    [switch]$AddActiveSyncDevices, 
    [switch]$Debug
)

#region HelperFunctions

#Function to get permissions
#enumerate groups: http://stackoverflow.com/questions/8055338/listing-users-in-ad-group-recursively-with-powershell-script-without-cmdlets/8055996#8055996
#ADSI Adapter: http://social.technet.microsoft.com/wiki/contents/articles/4231.working-with-active-directory-using-powershell-adsi-adapter.aspx
Function Get-Permissions(){
	param(
        [switch]$UseListOfUsers,
		[switch]$RemoveServiceAccts
    )
	
	if($RemoveServiceAccts){
		$listOfServiceAcctsToRemove = Read-Host "Specify path to Service Accounts CSV file (use 'email' header)..."
		If(-not (Test-Path $listOfServiceAcctsToRemove)){
            Write-LogEntry -LogName:$Script:LogFile -LogEntryText "$($listOfServiceAcctsToRemove) file not found." 
            throw "$($listOfServiceAcctsToRemove) file not found."
        }

        $svcAccts = Import-Csv $listOfServiceAcctsToRemove | Group Email -AsHashTable -AsString
        if($svcAccts -eq $null){
            Write-LogEntry -LogName:$Script:LogFile -LogEntryText "Problem with CSV file for Service Accounts, make sure to use 'email' as header."
            throw "Problem with CSV file for Service Accounts, make sure to use 'email' as header."
        }
	}

	If ($UseListOfUsers) { 
            $ImportFile = Read-Host "Specify path to Users CSV file  (use 'email' header)..."
            If(-not (Test-Path $ImportFile)){
                Write-LogEntry -LogName:$Script:LogFile -LogEntryText "$($ImportFile) file not found." 
                throw "$($ImportFile) file not found."
            }
            $mailboxesInCSV = Import-Csv $ImportFile | %{$_.Email}
            $Script:mailboxesinCSVHash = @{} 
            $mailboxesInCSV | %{$Script:mailboxesinCSVHash[$_] = $_ | out-null } #Used for quick search when separating permissioned users that are not part of the initial csv file
            
            if($mailboxesInCSV -eq $null){
                Write-LogEntry -LogName:$Script:LogFile -LogEntryText "Problem with CSV file for Users, make sure to use 'email' as header."
                throw "Problem with CSV file for Users, make sure to use 'email' as header."
            }
			Write-LogEntry -LogName:$Script:LogFile -LogEntryText "Number of users in csv file: $($mailboxesInCSV.count)" -ForegroundColor Gray
            $Set = $mailboxesInCSV | % {Get-Mailbox $_ -ResultSize Unlimited}
	}
    Else { 
            $Set = Get-Mailbox -ResultSize Unlimited | Where-Object {($_.WindowsEmailAddress -notlike "*@cookregentec.com") -and ($_.WindowsEmailAddress -notlike "*@cookmyosite.com") -and ($_.WindowsEmailAddress -notlike "*@medinstitute.com")}
	}

    #Variables
    $gatherfullaccess = $false
    $gatherdelegates= $true
    $gathercalendar = $true
    $gathersendas = $true
    $permsOutput = New-Object -TypeName System.Text.StringBuilder
    "User,Delegate,Permissions" > $Script:PermissionsFile
	
    $dse = [ADSI]"LDAP://Rootdse"
    $ext = [ADSI]("LDAP://CN=Extended-Rights," + $dse.ConfigurationNamingContext)
    $dn = [ADSI]"LDAP://$($dse.DefaultNamingContext)"
    $dsLookFor = new-object System.DirectoryServices.DirectorySearcher($dn)
    $permission = "Send As"
    $right = $ext.psbase.Children | ? { $_.DisplayName -eq $permission }

	
    
    $mailboxCounter = 0
    $setSize = $set.count
    Foreach ($mailbox in $Set) {
		#Progress Activity
        if($setSize -gt 0){
            $mailboxCounter++
            Write-Progress -Activity "Step 1 of 3: Gathering Permissions" -status "Items processed: $($mailboxCounter) of $($setSize)" `
    		        -percentComplete (($mailboxCounter / $setSize)*100)
        }
        
        $hasDependentPermission = $false
		$Error.clear();
        $ID = $mailbox.PrimarySMTPAddress.ToString();

		Write-LogEntry -LogName:$Script:LogFile -LogEntryText "Collect permissions for $($ID)" -ForegroundColor Yellow
        $delegates = $null
        If ($gatherdelegates) {
			Write-LogEntry -LogName:$Script:LogFile -LogEntryText "+Get ALL GrantSendOnBehalfTo permissions" 
            $delegates = $mailbox.grantsendonbehalfto.ToArray()
        }

        $aclFullAccess = $null
        If ($gatherfullaccess) {
            $aclAll = Get-MailboxPermission -Identity $mailbox -ErrorAction SilentlyContinue | ? {($_.AccessRights -like "*FullAccess*") -and ($_.IsInherited -eq $false) -and -not ($_.User -like "NT AUTHORITY\SELF") -and -not ($_.User -like "S-1-5*")}
            $aclFullAccess = new-object -TypeName System.Collections.ArrayList
            $aclAll | %{$aclFullAccess.add($_.user)}

            If($aclFullAccess){
                #Group enumeration
                foreach($perm in $($aclFullAccess)){
                    #Enumerate Group and add permission
				    $ifGroup = Get-Group $perm.ToString() -ErrorAction SilentlyContinue; $Error.Clear()
                    If($ifGroup){
                        $dsLookFor.Filter = "(&(memberof:1.2.840.113556.1.4.1941:=$($ifGroup.distinguishedName))(objectCategory=user))" 
                        $dsLookFor.SearchScope = "subtree" 
                        $mail = $dsLookFor.PropertiesToLoad.Add("mail")
                        $lstUsr = $dsLookFor.findall()
                        foreach ($usrTmp in $lstUsr) {
                            $usrTmp.Properties["mail"] | %{$aclFullAccess.add($_)}
                        }
                        $aclFullAccess.remove($perm)
                    }
                }
            }
        }

        $acl = $null
        If ($gathercalendar) {
            Write-LogEntry -LogName:$Script:LogFile -LogEntryText "+Get ALL Calendar-Folder permissions" 
			#$Calendar = $Mailbox.PrimarySmtpAddress.ToString() + ":\Calendar"
			$Calendar = (($Mailbox.PrimarySmtpAddress.ToString())+ ":\" + (Get-MailboxFolderStatistics -Identity $Mailbox.DistinguishedName | where-object {$_.FolderType -eq "Calendar"} | Select-Object -First 1).Name)
            $aclAll = Get-MailboxFolderPermission -Identity $Calendar 
			$acl = new-object -TypeName System.Collections.ArrayList
            If($aclAll){
                $aclAll | %{$acl.add($_.user)}
            }

            If($acl){
                #Group enumeration
                foreach($perm in $($acl)){
                    #Enumerate Group and add permission
				    $ifGroup = Get-Group $perm.ToString() -ErrorAction SilentlyContinue; $Error.Clear()
                    If($ifGroup){
						$groupName = $ifGroup.ToString().Split("/")
						$groupName = $groupName[$groupName.Length - 1]
						If(-not ($excludedGroups -contains $groupName)){
	                        Write-LogEntry -LogName:$Script:LogFile -LogEntryText "+Group Found: $($ifGroup)" 
							$dsLookFor.Filter = "(&(memberof:1.2.840.113556.1.4.1941:=$($ifGroup.distinguishedName))(objectCategory=user))" 
	                        $dsLookFor.PageSize  = 1000
	                        $dsLookFor.SearchScope = "subtree" 
	                        $mail = $dsLookFor.PropertiesToLoad.Add("mail")
	                        $lstUsr = $dsLookFor.findall()
	                        foreach ($usrTmp in $lstUsr) {
	                            $usrTmp.Properties["mail"] | %{$acl.add($_)}
	                        }
	                        $acl.remove($perm)
						}
                    }
                }
            }
        }

        $saPermissions = $null
        If ($gathersendas) {
            Write-LogEntry -LogName:$Script:LogFile -LogEntryText "+Get ALL Send As permissions" 
			$userDN = [ADSI]("LDAP://$($mailbox.DistinguishedName)")
            $saPermissions = new-object -TypeName System.Collections.ArrayList
			
            #Do not include inherited permissions. Only explicit permissions are migrated https://technet.microsoft.com/en-us/library/jj200581(v=exchg.150).aspx
            $userDN.psbase.ObjectSecurity.Access | ? { ($_.ObjectType -eq [GUID]$right.RightsGuid.Value) -and ($_.IsInherited -eq $false) } | select -ExpandProperty identityreference | %{
				If(-not ($_ -like "NT AUTHORITY\SELF")){
					$saPermissions.add($_)
				}
			}

            #Include inherited permissions
            #$userDN.psbase.ObjectSecurity.Access | ? { $_.ObjectType -eq [GUID]$right.RightsGuid.Value } | select -ExpandProperty identityreference | %{
			#	If(-not ($_ -like "NT AUTHORITY\SELF")){
			#		$saPermissions.add($_)
			#	}
			#}
            
            If($saPermissions){
                foreach($perm in $($saPermissions)){
                    #Enumerate Group and add permission
				    $ifGroup = Get-Group $perm.ToString() -ErrorAction SilentlyContinue; $Error.Clear()
                    If($ifGroup){
						$groupName = $ifGroup.ToString().Split("/")
						$groupName = $groupName[$groupName.Length - 1]
						If(-not ($excludedGroups -contains $groupName)){
							Write-LogEntry -LogName:$Script:LogFile -LogEntryText "+Group Found: $($ifGroup)" 
	                        $dsLookFor.Filter = "(&(memberof:1.2.840.113556.1.4.1941:=$($ifGroup.distinguishedName))(objectCategory=user))" 
	                        $dsLookFor.PageSize = 1000 
                            $dsLookFor.SearchScope = "subtree" 
	                        $mail = $dsLookFor.PropertiesToLoad.Add("mail")
	                        $lstUsr = $dsLookFor.findall()
	                        foreach ($usrTmp in $lstUsr) {
	                          $usrTmp.Properties["mail"] | %{$saPermissions.add($_)}
	                        }
	                        $saPermissions.remove($perm)
						}
                    }
                }
            }
        }

        If ($Error -ne $null) {Write-LogEntry -LogName:$Script:LogFile -LogEntryText "$ID : $Error" }
        Else {
            If ($aclFullAccess -ne $null) {
                ForEach ($ace in $aclFullAccess) {
                    $Error.clear();      
                    $recipient = get-recipient -Identity $ace.ToString() -ErrorAction SilentlyContinue;
                        If ($Error -ne $null) {Write-LogEntry -LogName:$Script:LogFile -LogEntryText "$ID : $Error" }
                        Else {
	                        If($mailbox.primarySMTPAddress -and $recipient.primarySMTPAddress){
						        If(-not ($mailbox.primarySMTPAddress.ToString() -eq $recipient.primarySMTPAddress.ToString())){
							        If($RemoveServiceAccts){
								        if(-not ($svcAccts.contains($recipient.primarySMTPAddress.tostring()) -or $svcAccts.contains($mailbox.primarySMTPAddress.ToString()))){
									        Write-host "+Found Full Permission: $($mailbox.primarySMTPAddress.ToString()) for $($recipient.primarySMTPAddress)" 
                                            $permsOutput.AppendLine("$($mailbox.primarySMTPAddress.ToString()),$($recipient.primarySMTPAddress),Full Access") | Out-Null
                                            $hasDependentPermission = $true
								        }
							        }
							        Else{
								        Write-host "+Found Full Permission: $($mailbox.primarySMTPAddress.ToString()) for $($recipient.primarySMTPAddress)" 
								        $permsOutput.AppendLine("$($mailbox.primarySMTPAddress.ToString()),$($recipient.primarySMTPAddress),Full Access") | Out-Null
                                        $hasDependentPermission = $true
							        }
						        }
					        }
                        }
                }
            }


            If ($acl -ne $null) {
                Write-LogEntry -LogName:$Script:LogFile -LogEntryText "+Found $($acl.count) entries with Calendar-Folder Access. Keeping only those with mailboxes." 
				ForEach ($ace in $acl) {
                $Error.clear();      
                $recipient = get-recipient -Identity $ace.ToString() -ErrorAction SilentlyContinue;
                    If ($Error -ne $null) {If((-not $Error.categoryinfo.reason -like "*ManagementObjectNotFoundException*")){Write-LogEntry -LogName:$Script:LogFile -LogEntryText "$ID : $Error" -ForegroundColor Gray}}
                    Else {
	                    If($mailbox.primarySMTPAddress -and $recipient.primarySMTPAddress){
							If(-not ($mailbox.primarySMTPAddress.ToString() -eq $recipient.primarySMTPAddress.ToString())){
								If($RemoveServiceAccts){
									if(-not ($svcAccts.contains($recipient.primarySMTPAddress.tostring()) -or $svcAccts.contains($mailbox.primarySMTPAddress.ToString()))){
                                        Write-LogEntry -LogName:$Script:LogFile -LogEntryText "+Found Calendar-Folder Permission: $($mailbox.primarySMTPAddress.ToString()) for $($recipient.primarySMTPAddress)" -ForegroundColor Gray
                                        $permsOutput.AppendLine("$($mailbox.primarySMTPAddress.ToString()),$($recipient.primarySMTPAddress),Calendar") | Out-Null
                                        $hasDependentPermission = $true
									}
								}
								Else{
                                    Write-LogEntry -LogName:$Script:LogFile -LogEntryText "+Found Calendar-Folder Permission: $($mailbox.primarySMTPAddress.ToString()) for $($recipient.primarySMTPAddress)" -ForegroundColor Gray
									$permsOutput.AppendLine("$($mailbox.primarySMTPAddress.ToString()),$($recipient.primarySMTPAddress),Calendar") | Out-Null
                                    $hasDependentPermission = $true
								}
							}
						}
                    }
                }
            }
            If ($delegates -ne $null) {
				Write-LogEntry -LogName:$Script:LogFile -LogEntryText "+Found $($delegates.count) entries with Delegate Perms. Keeping only those with mailboxes." 
                ForEach ($delegate in $delegates) {
                $Error.clear();      
                $recipient = get-recipient -Identity $delegate -ErrorAction SilentlyContinue;
                    If ($Error -ne $null) {If((-not $Error.categoryinfo.reason -like "*ManagementObjectNotFoundException*")){Write-LogEntry -LogName:$Script:LogFile -LogEntryText "$ID : $Error" -ForegroundColor Gray}}
                    Else {
						If($mailbox.primarySMTPAddress -and $recipient.primarySMTPAddress){
							If(-not ($mailbox.primarySMTPAddress.ToString() -eq $recipient.primarySMTPAddress.ToString())){
								If($RemoveServiceAccts){
									if(-not ($svcAccts.contains($recipient.primarySMTPAddress.tostring()) -or $svcAccts.contains($mailbox.primarySMTPAddress.ToString()))){
                                        Write-LogEntry -LogName:$Script:LogFile -LogEntryText "+Found Delegate: $($mailbox.primarySMTPAddress.ToString()) for $($recipient.primarySMTPAddress)" -ForegroundColor Gray
										$permsOutput.AppendLine("$($mailbox.primarySMTPAddress.ToString()),$($recipient.primarySMTPAddress),Delegation") | Out-Null
                                        $hasDependentPermission = $true
									}
								}
								Else{
                                    Write-LogEntry -LogName:$Script:LogFile -LogEntryText "+Found Delegate: $($mailbox.primarySMTPAddress.ToString()) for $($recipient.primarySMTPAddress)" -ForegroundColor Gray
                                    $permsOutput.AppendLine("$($mailbox.primarySMTPAddress.ToString()),$($recipient.primarySMTPAddress),Delegation") | Out-Null
									$hasDependentPermission = $true
								}
							}
						}
                    }
                }
            }
            If ($saPermissions -ne $null) {
				Write-LogEntry -LogName:$Script:LogFile -LogEntryText "+Found $($saPermissions.count) entries with Send As perms. Keeping only those with mailboxes."
                ForEach ($saPermission in $saPermissions) {
                    $Error.clear();
                    $recipient = get-recipient -Identity $saPermission.ToString() -ErrorAction SilentlyContinue;
                        If ($Error -ne $null) {If((-not $Error.categoryinfo.reason -like "*ManagementObjectNotFoundException*")){Write-LogEntry -LogName:$Script:LogFile -LogEntryText "$ID : $Error" -ForegroundColor Gray}}
                        Else {
						    If($mailbox.primarySMTPAddress -and $recipient.primarySMTPAddress){
							    If(-not ($mailbox.primarySMTPAddress.ToString() -eq $recipient.primarySMTPAddress.ToString())){
								    If($RemoveServiceAccts){
									    if(-not ($svcAccts.contains($recipient.primarySMTPAddress.tostring()) -or $svcAccts.contains($mailbox.primarySMTPAddress.ToString()))){
                                            Write-LogEntry -LogName:$Script:LogFile -LogEntryText "+Found Send As Permission: $($mailbox.primarySMTPAddress.ToString()) for $($recipient.primarySMTPAddress)" -ForegroundColor Gray
                                            $permsOutput.AppendLine("$($mailbox.primarySMTPAddress.ToString()),$($recipient.primarySMTPAddress),Send As") | Out-Null
										    $hasDependentPermission = $true
									    }
								    }
								    Else{
                                        Write-LogEntry -LogName:$Script:LogFile -LogEntryText "+Found Send As Permission: $($mailbox.primarySMTPAddress.ToString()) for $($recipient.primarySMTPAddress)" -ForegroundColor Gray
                                        $permsOutput.AppendLine("$($mailbox.primarySMTPAddress.ToString()),$($recipient.primarySMTPAddress),Send As") | Out-Null
									    $hasDependentPermission = $true
								    }
							    }
						    }
                        }
                }
            }
			If($hasDependentPermission -ne $true){
				If($mailbox.primarySMTPAddress){
					If($RemoveServiceAccts){
						If($svcAccts.contains($mailbox.primarySMTPAddress.ToString()) -ne $true){
                            $permsOutput.AppendLine("$($mailbox.primarySMTPAddress.ToString()),None,None") | Out-Null
						}
					}
					Else{
                        $permsOutput.AppendLine("$($mailbox.primarySMTPAddress.ToString()),None,None") | Out-Null
					}
				}
			}
       }
    }
    Write-LogEntry -LogName:$Script:LogFile -LogEntryText "Writing out permissions to file" -ForegroundColor Gray
    $permsOutput.ToString().TrimEnd() >> $Script:PermissionsFile
}

#Helper function for logging
Function Write-LogEntry {
   param(
      [string] $LogName ,
      [string] $LogEntryText,
      [string] $ForegroundColor
   )
   if ($LogName -NotLike $Null) {
      # log the date and time in the text file along with the data passed
      "$([DateTime]::Now.ToShortDateString()) $([DateTime]::Now.ToShortTimeString()) : $LogEntryText" | Out-File -FilePath $LogName -append;
      if ($ForeGroundColor -NotLike $null) {
         # for testing i pass the ForegroundColor parameter to act as a switch to also write to the shell console
         write-host $LogEntryText -ForegroundColor $ForeGroundColor
      }
   }
}

#Function to create batches
Function Create-Batches(){
    param(
        [Parameter(Mandatory=$true)]
        [string]$InputPermissionsFile
    )
		
    #Variables
    If(-not (Test-Path $InputPermissionsFile)){
        throw [System.IO.FileNotFoundException] "$($InputPermissionsFile) file not found."
    }
    
    $data = import-csv $InputPermissionsFile
    $hashData = $data | Group User -AsHashTable -AsString
	$hashDataByDelegate = $data | Group Delegate -AsHashTable -AsString
	$usersWithNoDependents = New-Object System.Collections.ArrayList
    $batch = @{}
    $batchNum = 0
    $hashDataSize = $hashData.Count
    $yyyyMMdd = Get-Date -Format 'yyyyMMdd'
	
    try{
        Write-LogEntry -LogName:$Script:LogFile -LogEntryText "Start function Create-Batches" -ForegroundColor Gray
        Write-LogEntry -LogName:$Script:LogFile -LogEntryText "Build ArrayList for users with no dependents" -ForegroundColor Gray
        If($hashDataByDelegate["None"].count -gt 0){
		    $hashDataByDelegate["None"] | %{$_.user} | %{[void]$usersWithNoDependents.Add($_)}
	    }	    

        Write-LogEntry -LogName:$Script:LogFile -LogEntryText "Identify users with no permissions on them, nor them have perms on another" -ForegroundColor Gray
	    If($usersWithNoDependents.count -gt 0){
		    $($usersWithNoDependents) | %{
			    if($hashDataByDelegate.ContainsKey($_)){
				    $usersWithNoDependents.Remove($_)
			    }	
		    }
            
            Write-LogEntry -LogName:$Script:LogFile -LogEntryText "Remove users with no dependents from hash Data" -ForegroundColor Gray 
		    $usersWithNoDependents | %{$hashData.Remove($_)}
		    #Clean out hashData of users in hash data with no delegates, otherwise they'll get batched
            Write-LogEntry -LogName:$Script:LogFile -LogEntryText "Clean out hashData of users in hash data with no delegates" -ForegroundColor Gray 
		    foreach($key in $($hashData.keys)){
                    if(($hashData[$key] | select -expandproperty Delegate ) -eq "None"){
				    $hashData.Remove($key)
			    }
		    }
	    }
        #Execute batch functions
        If(($hashData.count -ne 0) -or ($usersWithNoDependents.count -ne 0)){
            Write-LogEntry -LogName:$Script:LogFile -LogEntryText "Run Find-Links function" -ForegroundColor Gray 
            while($hashData.count -ne 0){Find-Links $hashData} 
            Write-LogEntry -LogName:$Script:LogFile -LogEntryText "Run Create-BatchFile function" -ForegroundColor Gray
            Create-BatchFile $batch $usersWithNoDependents
        }
    }
    catch {
        Write-LogEntry -LogName:$Script:LogFile -LogEntryText "Error: $_" -ForegroundColor Red 
    }
}

#Function to create batch file
Function Create-BatchFile($batchResults,$usersWithNoDepsResults){
	try{
         Write-LogEntry -LogName:$Script:LogFile -LogEntryText "Start function Create-BatchFile" -ForegroundColor Gray
         "Batch,User" > $Script:BatchesFile
	     foreach($key in $batchResults.keys){
            $batchNum++
            $batchName = "BATCH-$batchNum"
		    $output = New-Object System.Collections.ArrayList
		    $($batch[$key]) | %{$output.add($_.user)}
		    $($batch[$key]) | %{$output.add($_.delegate)}
		    $output | select -Unique | % {
               "$batchName"+","+$_ >> $Script:BatchesFile
		    }
         }
	     If($usersWithNoDepsResults.count -gt 0){
		     $batchNum++
		     foreach($user in $usersWithNoDepsResults){
		 	    #$batchName = "BATCH-$batchNum"
                $batchName = "BATCH-NoPermsOrDependents"
			    "$batchName"+","+$user >> $Script:BatchesFile
		     }
	     }
         Write-LogEntry -LogName:$Script:LogFile -LogEntryText "Batches created: $($batchNum)" -ForegroundColor Gray 
         Write-LogEntry -LogName:$Script:LogFile -LogEntryText "Number of comparisons: $($Script:comparisonCounter)" -ForegroundColor Gray 
     }
     catch{
        Write-LogEntry -LogName:$Script:LogFile -LogEntryText "Error: $_" -ForegroundColor Red  
     }
} 

#Function to identify associations    
Function Find-Links($hashData){
    try{
        Write-LogEntry -LogName:$Script:LogFile -LogEntryText "Hash Data Size: $($hashData.count)" -ForegroundColor Gray
        $nextInHash = $hashData.Keys | select -first 1
        $batch.Add($nextInHash,$hashData[$nextInHash])
	
	    Do{
		    $checkForMatches = $false
		    foreach($key in $($hashData.keys)){
	            $Script:comparisonCounter++ 
			
			    Write-Progress -Activity "Step 2 of 3: Analyzing Data" -status "Items remaining: $($hashData.Count)" `
    		    -percentComplete (($hashDataSize-$hashData.Count) / $hashDataSize*100)
			
	            #Checks
			    $usersHashData = $($hashData[$key]) | %{$_.user}
                $usersBatch = $($batch[$nextInHash]) | %{$_.user}
                $delegatesHashData = $($hashData[$key]) | %{$_.delegate} 
			    $delegatesBatch = $($batch[$nextInHash]) | %{$_.delegate}

			    $ifMatchesHashUserToBatchUser = [bool]($usersHashData | ?{$usersBatch -contains $_})
			    $ifMatchesHashDelegToBatchDeleg = [bool]($delegatesHashData | ?{$delegatesBatch -contains $_})
			    $ifMatchesHashUserToBatchDelegate = [bool]($usersHashData | ?{$delegatesBatch -contains $_})
			    $ifMatchesHashDelegToBatchUser = [bool]($delegatesHashData | ?{$usersBatch -contains $_})
			
			    If($ifMatchesHashDelegToBatchDeleg -OR $ifMatchesHashDelegToBatchUser -OR $ifMatchesHashUserToBatchUser -OR $ifMatchesHashUserToBatchDelegate){
	                if(($key -ne $nextInHash)){ 
					    $batch[$nextInHash] += $hashData[$key]
					    $checkForMatches = $true
	                }
	                $hashData.Remove($key)
	            }
	        }
	    } Until ($checkForMatches -eq $false)
        
        return $hashData 
	}
	catch{
        Write-LogEntry -LogName:$Script:LogFile -LogEntryText "Error: $_" -ForegroundColor Red
    }
}

#Function to create a migration schedule
Function Create-MigrationSchedule(){
    param(
        [Parameter(Mandatory=$true)]
        [string]$InputBatchesFile, 
        [switch]$IncludeActiveSyncDevices
    )
	try{
        #Variables
        If(-not (Test-Path $InputBatchesFile)){
            throw [System.IO.FileNotFoundException] "$($InputBatchesFile) file not found."
        }
        $usersFromBatch = import-csv $InputBatchesFile
        If($IncludeActiveSyncDevices){
            "Migration Date(MM/dd/yyyy),Migration Window,Migration Group,PrimarySMTPAddress,SuggestedBatch,MailboxSize(MB),ActiveSyncDevices,Notes" > $Script:MigrationScheduleFile
        }
        Else{
            "Migration Date(MM/dd/yyyy),Migration Window,Migration Group,PrimarySMTPAddress,SuggestedBatch,MailboxSize(MB),Notes" > $Script:MigrationScheduleFile
        }
        $userInfo = New-Object System.Text.StringBuilder
        Write-LogEntry -LogName:$Script:LogFile -LogEntryText "Number of users in the migration schedule: $($usersFromBatch.Count)" -ForegroundColor Gray

        $usersFromBatchCounter = 0
        foreach($item in $usersFromBatch){
            $usersFromBatchCounter++
            $usersFromBatchRemaining = $usersFromBatch.count - $usersFromBatchCounter
            Write-Progress -Activity "Step 3 of 3: Creating migration schedule" -status "Items remaining: $($usersFromBatchRemaining)" `
    		    -percentComplete (($usersFromBatchCounter / $usersFromBatch.count)*100)

           #Check if using UseImportCSVFile and if yes, check if the user was part of that file, otherwise mark 
           $isUserPartOfInitialCSVFile = ""
           If($Script:mailboxesinCSVHash){ 
            If(-not $Script:mailboxesinCSVHash.ContainsKey($item.user)){
                $isUserPartOfInitialCSVFile = "User was not part of initial csv file"
            }
           }
           $user = get-user $item.User -erroraction SilentlyContinue
		   
           If(![string]::IsNullOrEmpty($user.WindowsEmailAddress)){
			$mbStats = Get-MailboxStatistics $user.WindowsEmailAddress.tostring() 
			If($mbStats){
				$mailboxSize = ($mbStats).TotalItemSize.Value.ToMb()
			}
			Else{
                $mailboxSize = 0
            }

            If(-not $IncludeActiveSyncDevices){
                $userInfo.AppendLine(",,,$($user.WindowsEmailAddress),$($item.Batch),$($mailboxSize),$isUserPartOfInitialCSVFile") | Out-Null
            }
            Else{
                #check for active sync devices
                $activeSyncInfo = Get-CASMailbox $user.WindowsEmailAddress.tostring()
                #If(($activeSyncInfo.hasactivesyncdevicepartnership -eq $true) -AND ($activeSyncInfo.activesyncalloweddeviceIDs -ne $null)){
                If($activeSyncInfo.activesyncalloweddeviceIDs -ne $null){
                    $activeSyncDevices = New-Object System.Text.StringBuilder
                    $activeSyncInfo.activesyncalloweddeviceIDs | % {$activeSyncDevices.append("|$_")} | Out-Null
                    $userInfo.AppendLine(",,,$($user.WindowsEmailAddress),$($item.Batch),$($mailboxSize),$($activeSyncDevices.ToString()),$isUserPartOfInitialCSVFile") | Out-Null
                }
                else{
                    $userInfo.AppendLine(",,,$($user.WindowsEmailAddress),$($item.Batch),$($mailboxSize),,$isUserPartOfInitialCSVFile") | Out-Null
                }

            }
           }
		   Else{ #there was an error either getting the user from Get-User or the user doesn't have an email address
			   	If(-not $IncludeActiveSyncDevices){
	                $userInfo.AppendLine(",,,$($item.User),$($item.Batch),n/a,User not found or doesn't have an email address") | Out-Null
	            }
	            Else{
					$userInfo.AppendLine(",,,$($item.User),$($item.Batch),n/a,,User not found or doesn't have an email address") | Out-Null
				}
		   }
        }
        $userInfo.ToString().TrimEnd() >> $Script:MigrationScheduleFile
    }
    catch{
        Write-LogEntry -LogName:$Script:LogFile -LogEntryText "Error: $_" -ForegroundColor Red
    }
}

#endregion HelperFunctions

#Main

$timeTaken = (Measure-Command { 
    Write-LogEntry -LogName:$Script:LogFile -LogEntryText "START Script..." -ForegroundColor Yellow

    #Get Exchange Version, Add Snapin, and ViewEntireForest
    #Build numbers: https://technet.microsoft.com/en-us/library/hh135098(v=exchg.150).aspx
    $ExchangeVersion = GCM Exsetup.exe | % {$_.FileVersionInfo} | select -ExpandProperty FileVersion
    $Build = ($ExchangeVersion.tostring()).Split(".")[0]

    If(($Build -eq "08") -or ($Build -eq "8")){Add-PsSnapin Microsoft.Exchange.Management.PowerShell.Admin -ErrorAction SilentlyContinue; $AdminSessionADSettings.ViewEntireForest = $True}
    ElseIf($Build -eq "14"){Add-PsSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue; Set-AdServerSettings -ViewEntireForest $True}
    ElseIf($Build -eq "15"){Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue; Set-AdServerSettings -ViewEntireForest $True}

    $invocationStatus = $MyInvocation.MyCommand.Path
    If($invocationStatus){
        #Use for older powershell versions
        $scriptPath = split-path $MyInvocation.MyCommand.Path -parent
    }
    ElseIf($PSScriptRoot){
        #Powershell v3+
        $scriptPath = $PSScriptRoot
    }
    Else{
        Write-LogEntry -LogName:$LogFile -LogEntryText "Unable to get directory of script, saving logs to c:\" -ForegroundColor Yellow
        $scriptPath = "c:"
    }
    
    $Script:comparisonCounter = 0
    $yyyyMMdd = Get-Date -Format 'yyyyMMdd'
    $Script:LogFile = "$scriptPath\FTC-MigrationPlanningLogFile-$yyyyMMdd.txt"
    $Script:PermissionsFile = "$scriptPath\FTC-PermissionsOutput.csv"
    $Script:BatchesFile = "$scriptPath\FTC-BatchesOutput.csv"
    $Script:MigrationScheduleFile = "$scriptPath\FTC-MigrationSchedule.csv"
    $Script:excludedGroups = "Exchange Organization Administrators","Organization Management","Exchange Servers"

    If($Debug){
        Write-LogEntry -LogName:$LogFile -LogEntryText "DEBUG MODE..." -ForegroundColor Yellow
        #measure-command{ . '.\FTCMigrationPlanning.ps1'}
	    #Get-Permissions 
        Create-Batches -InputPermissionsFile $Script:PermissionsFile
        Create-MigrationSchedule -InputBatchesFile $Script:BatchesFile
        exit
    }

    #Execute Functions
    if($UseImportCSV -and -not $ListActiveSyncDevices){
	    if($RemoveServiceAccts){
            Write-LogEntry -LogName:$LogFile -LogEntryText "STEP 1: Gathering Permissions..." -ForegroundColor Yellow
            Get-Permissions -UseListOfUsers -RemoveServiceAccts 

            Write-LogEntry -LogName:$LogFile -LogEntryText "STEP 2: Creating batches..." -ForegroundColor Yellow
            Create-Batches -InputPermissionsFile $Script:PermissionsFile
        
            Write-LogEntry -LogName:$LogFile -LogEntryText "STEP 3: Creating migration schedule..." -ForegroundColor Yellow
            Create-MigrationSchedule -InputBatchesFile $Script:BatchesFile
	    }
	    else{
            Write-LogEntry -LogName:$LogFile -LogEntryText "STEP 1: Gathering Permissions..." -ForegroundColor Yellow
	        Get-Permissions -UseListOfUsers 

            Write-LogEntry -LogName:$LogFile -LogEntryText "STEP 2: Creating batches..." -ForegroundColor Yellow
            Create-Batches -InputPermissionsFile $Script:PermissionsFile

            Write-LogEntry -LogName:$LogFile -LogEntryText "STEP 3: Creating migration schedule..." -ForegroundColor Yellow
            Create-MigrationSchedule -InputBatchesFile $Script:BatchesFile
	    }
    }
    elseif($UseImportCSV -and $ListActiveSyncDevices){
        if($RemoveServiceAccts){
			 Write-LogEntry -LogName:$LogFile -LogEntryText "STEP 1: Gathering Permissions..." -ForegroundColor Yellow
            Get-Permissions -UseListOfUsers -RemoveServiceAccts 

            Write-LogEntry -LogName:$LogFile -LogEntryText "STEP 2: Creating batches..." -ForegroundColor Yellow
            Create-Batches -InputPermissionsFile $Script:PermissionsFile
        
            Write-LogEntry -LogName:$LogFile -LogEntryText "STEP 3: Creating migration schedule..." -ForegroundColor Yellow
            Create-MigrationSchedule -InputBatchesFile $Script:BatchesFile -IncludeActiveSyncDevices
	    }
	    else{
            Write-LogEntry -LogName:$LogFile -LogEntryText "STEP 1: Gathering Permissions..." -ForegroundColor Yellow
	        Get-Permissions -UseListOfUsers 

            Write-LogEntry -LogName:$LogFile -LogEntryText "STEP 2: Creating batches..." -ForegroundColor Yellow
            Create-Batches -InputPermissionsFile $Script:PermissionsFile

            Write-LogEntry -LogName:$LogFile -LogEntryText "STEP 3: Creating migration schedule..." -ForegroundColor Yellow
            Create-MigrationSchedule -InputBatchesFile $Script:BatchesFile -IncludeActiveSyncDevices
	    }
    }
    elseif($ListActiveSyncDevices){
        if($RemoveServiceAccts){
		    Write-LogEntry -LogName:$LogFile -LogEntryText "STEP 1: Gathering Permissions..." -ForegroundColor Yellow
	        Get-Permissions -RemoveServiceAccts 

            Write-LogEntry -LogName:$LogFile -LogEntryText "STEP 2: Creating batches..." -ForegroundColor Yellow
            Create-Batches -InputPermissionsFile $Script:PermissionsFile
        
            Write-LogEntry -LogName:$LogFile -LogEntryText "STEP 3: Creating migration schedule..." -ForegroundColor Yellow
            Create-MigrationSchedule -InputBatchesFile $Script:BatchesFile -IncludeActiveSyncDevices
	    }
	    else{
		    Write-LogEntry -LogName:$LogFile -LogEntryText "STEP 1: Gathering Permissions..." -ForegroundColor Yellow
	        Get-Permissions 
            
            Write-LogEntry -LogName:$LogFile -LogEntryText "STEP 2: Creating batches..." -ForegroundColor Yellow
            Create-Batches -InputPermissionsFile $Script:PermissionsFile
        
            Write-LogEntry -LogName:$LogFile -LogEntryText "STEP 3: Creating migration schedule..." -ForegroundColor Yellow
            Create-MigrationSchedule -InputBatchesFile $Script:BatchesFile -IncludeActiveSyncDevices
	    }
    }
    ElseIf($AddActiveSyncDevices){
        Write-LogEntry -LogName:$Script:LogFile -LogEntryText "Connecting to Exchange Online..." -ForegroundColor Yellow
        $credential = get-credential
        $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $credential -Authentication "Basic" -AllowRedirection
        Import-PSSession $exchangeSession -AllowClobber -DisableNameChecking
       
        $schedule = Import-Csv $Script:MigrationScheduleFile | group primarysmtpaddress -AsHashTable -asstring
        
        $migratedUsers = Get-Mailbox | ?{$_.MailboxMoveStatus -eq "completed"}  
        Write-LogEntry -LogName:$Script:LogFile -LogEntryText "Please give us a few seconds while we find migrated users to add previous mobile device registrations:" -ForegroundColor Yellow

        #Migrated users with no activesyncdevices added
        $migratedUsersWOActiveSyncDevices = $migratedUsers | Get-CASMailbox | %{If(-not $_.activesyncalloweddeviceids){$_}}

        #Add ActiveSyncDevices for migrated users
        foreach($user in $migratedUsersWOActiveSyncDevices){
               $checkDevices = $schedule[$user.primarysmtpaddress.tostring()] | %{$_.ActiveSyncDevices}
               If($checkDevices){
                      $devices = $checkDevices.split("|")
                      $devices = $devices | ?{$_ -ne ""}
                      Write-LogEntry -LogName:$Script:LogFile -LogEntryText "We found devices for migrated user: $($user.primarysmtpaddress). Adding this device into Exchange Online." -ForegroundColor Gray
                      Set-CASMailbox -identity $user.primarysmtpaddress.tostring() -ActiveSyncAllowedDeviceIDs @{Add=$devices} 
               }
               else{
                      Write-LogEntry -LogName:$Script:LogFile -LogEntryText "We didn’t find devices for migrated user: $($user.primarysmtpaddress). This may be expected, but please confirm. " -ForegroundColor Gray
               }
        }

        #remove exchange session once done
        Remove-PSSession $exchangeSession

        #Write output, finish script
        Write-LogEntry -LogName:$Script:LogFile -LogEntryText "Finished adding active sync devices for the migrated users. The following command can be used for verification: " -ForegroundColor Yellow
        Write-LogEntry -LogName:$Script:LogFile -LogEntryText "get-casmailbox <UserSMTPAddress> | select *device*" -ForegroundColor Yellow
        exit       
    }
    else{ #Default run with no parameters, collect permissions for all users
	    if($RemoveServiceAccts){
		    Write-LogEntry -LogName:$LogFile -LogEntryText "STEP 1: Gathering Permissions..." -ForegroundColor Yellow
	        Get-Permissions -RemoveServiceAccts 

            Write-LogEntry -LogName:$LogFile -LogEntryText "STEP 2: Creating batches..." -ForegroundColor Yellow
            Create-Batches -InputPermissionsFile $Script:PermissionsFile
        
            Write-LogEntry -LogName:$LogFile -LogEntryText "STEP 3: Creating migration schedule..." -ForegroundColor Yellow
            Create-MigrationSchedule -InputBatchesFile $Script:BatchesFile
	    }
	    else{
		    Write-LogEntry -LogName:$LogFile -LogEntryText "STEP 1: Gathering Permissions..." -ForegroundColor Yellow
	        Get-Permissions 
            
            Write-LogEntry -LogName:$LogFile -LogEntryText "STEP 2: Creating batches..." -ForegroundColor Yellow
            Create-Batches -InputPermissionsFile $Script:PermissionsFile
        
            Write-LogEntry -LogName:$LogFile -LogEntryText "STEP 3: Creating migration schedule..." -ForegroundColor Yellow
            Create-MigrationSchedule -InputBatchesFile $Script:BatchesFile
	    }
    }
}).TotalSeconds

Write-LogEntry -LogName:$LogFile -LogEntryText "Time taken in seconds: $($timeTaken)" -ForegroundColor Gray
Write-LogEntry -LogName:$Script:LogFile -LogEntryText "Migration schedule file location: $Script:MigrationScheduleFile" -ForegroundColor Yellow

     
