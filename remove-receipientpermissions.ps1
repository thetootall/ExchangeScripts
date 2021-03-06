#CSV will be formatted with one column of header "email" with addresses of Shared mailboxes
$file = Read-host "type file path (if not in current folder)"
$list = import-csv $file
Foreach ($m in $list){
#start L1
Write-host $m
$m = $m.email
[string]$mailboxID = Get-Mailbox -identity $m
#Edit - you dont need the following for Exchange on-premise
$mailboxID1 = $mailboxID.ExchangeGUID
$mailboxID1 = $mailboxID1.guid
write-host "Clearing $mailboxID"
#Exchange Online
$sendperms = Get-RecipientPermission -Identity $mailboxID1 | ?{($_.Identity -eq $MailboxID) -and ($_.IsInherited -eq $False) -and -not ($_.Trustee -match "NT AUTHORITY")} | Select *
#Exchange on-premise
#$sendperms = Get-ADPermission -Identity $mailboxID | ?{($_.Identity -eq $MailboxID) -and ($_.IsInherited -eq $False) -and -not ($_.User -match "NT AUTHORITY")} | Select *
$sendperms | fl
$agree = Read-host "Press Y to Continue"
If ($agree -eq "Y"){
#start L2
Foreach ($i in $sendperms){
#start L3
$listuser = $i.trustee 
$listidentity = $i.identity
write-host "removing trustee $listuser"
#Exchange Online
Remove-RecipientPermission -Identity $listidentity -Trustee $listuser -AccessRights SendAs -confirm:$false
#Exchange On-premise
#Remove-ADPermission -Identity $listidentity -TUser $listuser -ExtendedRights "Send As" -confirm:$false
#end L3
}
#end L2
}
#end L1
}
