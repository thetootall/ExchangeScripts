#CSV will be formatted with two column headers: "email" with addresses of Shared mailboxes, and "group" with the name of the group
#$file = Read-host "type file path (if not in current folder)"
$csvlist = import-csv "grouplist.csv"
ForEach ($item in $csvlist) {
#start L1
$e = $item.email
$g = $item.group
write-host "Adding permission $e adding $g"
$mailboxID = Get-EXOMailbox -PrimarySmtpAddress $e
$mailboxID1 = $mailboxID.Identity
$agree = Read-host "Press Y to Continue"
If ($agree -eq "Y"){
#start L2
Add-MailboxPermission -Identity $mailboxID1 -User $g -AccessRights FullAccess -confirm:$false
#end L2
}
#end L1
}