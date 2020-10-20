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
Get-MigrationUser –BatchID  $batchname | Get-MoveRequestStatistics | Select MailboxIdentity, DistinguishedName, Alias, ExchangeGuid, Status, StatusDetail, SyncStage, SourceServer, RemoteDatabaseName, BadItemsEncountered, LargeItemsEncountered, MissingItemsEncountered, QueuedTimestamp, StartTimestamp, LastUpdateTimestamp, LastSuccessfulSyncTimestamp, OverallDuration, TotalInProgressDuration, @{label="TotalMailboxSize(MB)";expression={$_.TotalMailboxSize.Value.ToMB()}}, TotalMailboxItemCount, @{label="BytesTransferred.(MB)";expression={$_.BytesTransferred.Value.ToMB()}}, ItemsTransferred, @{label="ShardBytesTransferred.(MB)";expression={$_.ShardBytesTransferred.Value.ToMB()}}, ShardItemsTransferred | export-csv RemoteOnBorading1Stats.csv -notype
