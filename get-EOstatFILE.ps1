
$logfile = "_MRLL_MAILBOXSTATS_REPORT2.txt"
#$Users = Get-Mailbox -resultsize unlimited
$Users = import-csv failedusers.csv

$header = "Mail|Name|Type|LastLogon|LastLogoff|ItemSize|ItemCount|DeletedItemSize|DeletedItemCount"
$header | Out-file $logfile -Append

Foreach ($item in $Users){

$myuser = $item.UserPrincipalName
$mymailbox = Get-mailbox $myuser

$mymail = $mymailbox.PrimarySMTPAddress
$mymailid = $mymailbox.DistinguishedName
$mymailtype = $mymailbox.RecipientType

$mymailboxstat = Get-MailboxStatistics $mymailid | select userprincipalname,displayname,RecipientType,lastlogontime,lastlogofftime,ItemCount,TotalItemSize,TotalDeletedItemSize,DeletedItemCount

$mymaildisp = $mymailboxstat.displayname
$mymaillogon = $mymailboxstat.lastlogontime
$mymaillogoff = $mymailboxstat.lastlogofftime
$mymailsizeitem = $mymailboxstat.ItemCount
$mymailsizedeleted = $mymailboxstat.DeletedItemCount

$mymailitem = [math]::Round(($mymailboxstat.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB),2)
$mymaildeleted = [math]::Round(($mymailboxstat.TotalDeletedItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB),2)

$output = "$mymail|$mymaildisp|$mymailtype|$mymaillogon|$mymaillogoff|$mymailitem|$mymailsizeitem|$mymaildeleted|$mymailsizedeleted"

Write-host $Output
$output | Out-file $logfile -Append
}