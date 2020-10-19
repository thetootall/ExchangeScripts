#New-managementroleassignment –Role "ApplicationImpersonation" –user sathesh@careexchange.onmicrosoft.com

$userList = Get-Mailbox | select identity,primarysmtpaddress
$update = "N"
$currtime = get-date -format yyyy-MM-dd_HH-mm-ss
$logfile = "personalfolders-$currtime.txt"
$header = "mailbox,status,policy"
$header | Out-file $logfile

foreach ($user in $userlist){
#start loop
$arr = $user.PrimarySMTPaddress
$arr1 = Get-MailboxFolderStatistics $arr |?{$_.folderpath -eq "/Personal Folders"}
If ($arr1 -eq $Null){
#New-MailboxFolder 
$msg = "$arr, path does not exist"
Write-host $msg  -ForegroundColor Red
$msg | out-file $logfile -append
if ($update -eq "Y"){
Write-host "Creating Personal Folders" -ForegroundColor Green
New-MailboxFolder -Parent $arr -Name "Personal Folders"}
}
Else
{
$arr2 = $arr1.identity
$arrdel = $arr1.deletepolicy
$msg2 =  "$arr, $arr2, $arrdel"
Write-host $msg2 -ForegroundColor Green
$msg2 | out-file $logfile -append
}
#endloop
}