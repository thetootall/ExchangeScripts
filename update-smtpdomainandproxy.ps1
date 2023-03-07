#update this path for the log file or it will use the directory the script is executed in
$logfile = "log.txt"
#update the SearchBase to target the location of the OU you want to process
$f = get-aduser -Filter * -SearchBase "OU=Company,DC=home,DC=local"

#loop through each item in the array
ForEach ($ff in $f) {
Write-host $ff
$thistime = get-date -Format MM-dd-yy-HH:ss
$thisuser = Get-ADuser $ff -Properties samaccountname, userprincipalname, targetAddress, mail, sn, givenName, proxyaddresses | select-object samaccountname,userprincipalname, targetAddress, mail, sn, givenName, @{n = "proxyAddress"; e = { $_.proxyAddresses | Where-object { $_ -clike "SMTP:*" } } } , mailNickname
#samaccountname is important as it is how get-aduser resolved users in the array
$thissam = $thisuser.samAccountName
$thisupn = $thisuser.UserPrincipalName
$thismail = $thisuser.mail
$thisproxy = $thisuser.proxyaddress
#create new nickname based on mailnickname field - this is recommended
#$thisnick = $thisuser.mailnickname
#$thischange = $thisnick + "@domain.com"
#createnickname based on first part of existing primary address - this is a common approach
$thisnick = $thismail.Split("@")[0] + "@domain.com"
$thismsg = "$thistime $thisupn will be changed from $thismail to $thisnick | $thisproxy"
Write-host $thismsg
$thismsg | out-file $logfile -append
          
#update the addresses
Set-ADUser -Identity $thissam -Remove @{Proxyaddresses=$thisproxy}
Set-ADUser -Identity $thissam -Add @{Proxyaddresses="SMTP:"+$thisnick}
Set-ADUser -Identity $thissam -Add @{Proxyaddresses="smtp:"+$thismail}
Set-ADUser -Identity $thissam -clear mail
Set-ADUser -Identity $thissam -Add @{mail="$thisnick"}

#reread AD and report updates
$updateuser = Get-ADuser $ff -Properties targetAddress, mail, sn, givenName, proxyaddresses | select-object targetAddress, mail, sn, givenName, @{n = "proxyAddress"; e = { $_.proxyAddresses | Where-object { $_ -clike "SMTP:*" } } } , mailNickname
$updatetime = get-date -Format MM-dd-yy-HH:ss
$updateproxy = $updateuser.proxyaddress
$updatemail = $updateuser.mail
$updatemsg = "$updatetime $thisupn now reads $updatemail | $updateproxy"
Write-host $updatemsg
$updatemsg | out-file $logfile -append
}
