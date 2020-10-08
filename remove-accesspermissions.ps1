#$file = Read-host "type file path (if not in current folder)"
$csvlist = import-csv "list.csv"
Foreach ($m in $csvlist){
#start L1
Write-host $m
$m = $m.email
$mailboxID = Get-Mailbox -identity $m
$mailboxID1 = $mailboxID.Identity
write-host "Clearing $mailboxID for $mailboxID1"
$fullperms = Get-MailboxPermission $mailboxID1 | ?{($_.Identity -eq $MailboxID) -and ($_.IsInherited -eq $False) -and -not ($_.User -match "NT AUTHORITY")} | Select *
$fullperms | fl
$agree = Read-host "Press Y to Continue"
If ($agree -eq "Y"){
#start L2
Foreach ($i in $fullperms){
#start L3
$listuser = $i.user 
$listidentity = $i.identity
write-host "removing user $listuser"
Remove-MailboxPermission -Identity $listidentity -User $listuser -AccessRights FullAccess -confirm:$false
#end L3
}
#end L2
}
#end L1
}
