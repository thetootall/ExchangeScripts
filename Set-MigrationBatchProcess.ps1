#http://www.networknet.nl/apps/wp/archives/2932
#https://docs.microsoft.com/en-us/powershell/module/exchange/new-migrationbatch?view=exchange-ps
#be sure to structure your CSV with a column header "EmailAddress" and 
$csvname = Read-Host "Enter path to CSV"

#get the OrdID = this is also the TargetDeliveryDomain
$orgID = (Get-OrganizationConfig).identity

#list of email addresses to contact during migration process
$emailadmins "user@client.com,admin@client.com"

#Populate the migration endpoint. NOTE If there are more than one 1 you will need to adjust
$migend = Get-MigrationEndpoint | ?{$_.endpointtype -like "ExchangeRemoteMove"}

#Create the Migration Batch
$batchname = Read-Host "Enter the migration batch name"
$OnboardingBatch = New-MigrationBatch -Name $batchname -SourceEndpoint $midend -TargetDeliveryDomain $orgID -BadItemLimit 10 -LargeItemLimit 5 -Autostart -NotificationEmails $emailadmins -CSVData ([System.IO.File]::ReadAllBytes($csvname)); 
Start-MigrationBatch -Identity $OnboardingBatch.Identity

#Run the following to output migration batch status
#Exchange Online acts differently with math; see https://4sysops.com/archives/sort-exchange-and-office-365-mailboxes-by-size-with-powershell/
Get-MigrationUser –BatchID  $batchname | Get-MoveRequestStatistics | Select MailboxIdentity, DistinguishedName, Alias, ExchangeGuid, Status, StatusDetail, SyncStage, SourceServer, RemoteDatabaseName, BadItemsEncountered, LargeItemsEncountered, MissingItemsEncountered, QueuedTimestamp, StartTimestamp, LastUpdateTimestamp, LastSuccessfulSyncTimestamp, OverallDuration, TotalInProgressDuration, @{name="TotalMailboxSize (MB)"; expression=[math]::Round(($_.TotalMailboxSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB),2)}}, TotalMailboxItemCount, @{name="BytesTransferred (MB)"; expression=[math]::Round(($_.BytesTransferred.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB),2)}}, ItemsTransferred, @{name="ShardBytesTransferred (MB)"; expression=[math]::Round(($_.ShardBytesTransferred.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB),2)}}, ShardItemsTransferred | export-csv RemoteOnBorading1Stats.csv -notype
