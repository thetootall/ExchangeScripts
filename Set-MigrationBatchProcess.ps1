#http://www.networknet.nl/apps/wp/archives/2932
#https://docs.microsoft.com/en-us/powershell/module/exchange/new-migrationbatch?view=exchange-ps
#be sure to structure your CSV with a column header "EmailAddress"
$emailadmins "user@client.com,admin@client.com"

#find the OrdID
$orgID = (Get-OrganizationConfig).identity

#Populate the migration endpoint. NOTE If there are more than one 1 you will need to adjust
$batchname = Read-Host "Enter the migration batch name"
$migend = Get-MigrationEndpoint | ?{$_.endpointtype -like "ExchangeRemoteMove"}

#Create the Migration Batch
$OnboardingBatch = New-MigrationBatch -Name $batchname -SourceEndpoint $midend -TargetDeliveryDomain $orgID -BadItemLimit 10 -LargeItemLimit 5 -Autostart -NotificationEmails $emailadmins -CSVData ([System.IO.File]::ReadAllBytes("C:\Users\Administrator\Desktop\RemoteOnBoarding1.csv")); 
Start-MigrationBatch -Identity $OnboardingBatch.Identity

#Run the following to output migration batch status
Get-MigrationUser –BatchID  $batchname | Get-MoveRequestStatistics | Select MailboxIdentity, DistinguishedName, Alias, ExchangeGuid, Status, StatusDetail, SyncStage, SourceServer, RemoteDatabaseName, BadItemsEncountered, LargeItemsEncountered, MissingItemsEncountered, QueuedTimestamp, StartTimestamp, LastUpdateTimestamp, LastSuccessfulSyncTimestamp, OverallDuration, TotalInProgressDuration, @{Expression={$_.TotalMailboxSize.Value.ToMB()}; Label="TotalMailboxSize"}, TotalMailboxItemCount,  @{Expression={$_.BytesTransferred.Value.ToMB()}; Label="BytesTransferred"}, ItemsTransferred, @{Expression={$_.ShardBytesTransferred.Value.ToMB()}; Label="ShardBytesTransferred"}, ShardItemsTransferred | export-csv RemoteOnBorading1Stats.csv -notype
