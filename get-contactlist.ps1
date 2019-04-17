
$update = "NO"
$logfile = "ADMINSYSlist.txt"
$header = "Object|Name|Email|ID|State"
$header | Out-file $logfile
$contactmatch = Get-ADObject -LDAPFilter "(&(objectclass=contact))" -SearchBase "OU=Power365,OU=Contacts,OU=ENT,DC=adminsys,DC=mrll,DC=com" -Properties * | ?{$_.mail -notlike "*mimdev.*"}

#$users = Import-Csv "TMpower365list1.csv"
ForEach ($Item in $contactmatch){

#$upn = $item.UserPrincipalName
$contactuser = $item

#$contactuser = $contactmatch | ?{$_.displayname -eq $usermatch}
$contactid = $contactuser.ObjectGUID
$contactdisp = $contactuser.DisplayName
$contacteml = $contactuser.mail
clear-variable contactstat
$contactstat = $contactuser.msExchHideFromAddressLists
If ($contactstat -eq $NULL){$contactstat = "NULL"}
$outputcontactdetail = "Loading Contact|$contactdisp|$contacteml|$contactid|$contactstat"
$outputcontactdetail | Out-file $logfile -Append
Write-host $outputcontactdetail

If ($update -eq "YES"){
Get-ADObject $contactid | set-adobject -Replace @{msExchHideFromAddressLists="TRUE"}
clear-variable contactupd
$contactupd = (Get-ADObject $contactid -Properties msExchHideFromAddressLists).msExchHideFromAddressLists
If ($contactupd -eq $NULL){$contactupd = "NULL"}
$outputcontactupdate = "Updated Contact Exchange attribute|$contacteml|$contactID|$contactupd"
$outputcontactupdate | Out-file $logfile -Append
Write-host $outputcontactupdate}

$myuser = Get-AdUser -filter {displayname -eq $contactdisp} -properties givenname,sn,displayname,mail,proxyaddresses,msExchHideFromAddressLists | ?{$_.mail -notlike "*mimdev.*"}
$usermatch = $myuser.DisplayName
$usersam = $myuser.SamAccountName
$userupn = $myuser.UserPrincipalName
$userstat = $myuser.msExchHideFromAddressLists
clear-variable userstat
If ($userstat -eq $NULL){$userstat = "NULL"}
$outputuserdetail = "Loading User|$usermatch|$userupn|$usersam|$userstat"
$outputuserdetail | Out-file $logfile -Append
Write-host $outputuserdetail

If ($update -eq "YES"){
Get-ADUser $usersam | set-aduser -Replace @{msExchHideFromAddressLists="FALSE"}
clear-variable userupd
$userupd = (Get-Aduser $usersam -Properties msExchHideFromAddressLists).msExchHideFromAddressLists
If ($userupd -eq $NULL){$userupd = "NULL"}
$outputuserupdate = "Updated User Exchange attribute|$userupn|$usersam|$userupd"
$outputuserupdate | Out-file $logfile -Append
Write-host $outputuserupdate

$outputdivider = "-----------------------------" 
$outputdivider | Out-file $logfile -Append
Write-host $outputdivider -ForegroundColor Red}

}