
$f = get-aduser -Filter * -SearchBase "OU=Company,DC=home,DC=local"

ForEach ($ff in $f) {
    Write-host $ff
        $thistime = get-date -Format MM-dd-yy-HH:ss
        $thisuser = Get-ADuser $ff -Properties samaccountname, userprincipalname, targetAddress, mail, sn, givenName, proxyaddresses | select-object samaccountname,userprincipalname, targetAddress, mail, sn, givenName, @{n = "proxyAddress"; e = { $_.proxyAddresses | Where-object { $_ -clike "SMTP:*" } } } , mailNickname
        $thissam = $thisuser.samAccountName
        $thisupn = $thisuser.UserPrincipalName
        $thismail = $thisuser.mail
        $thisproxy = $thisuser.proxyaddress
        #create new nickname based on mailnickname field
        #$thisnick = $thisuser.mailnickname
        #$thischange = $thisnick + "@circana.com"
        #createnickname based on first part of existing primary address
        $thisnick = $thismail.Split("@")[0] + "@circana.com"
        $thismsg = "$thistime $thisupn will be changed from $thismail to $thisnick | $thisproxy"
            Write-host $thismsg
            $thismsg | out-file log.txt -append
            

Set-ADUser -Identity $thissam -Remove @{Proxyaddresses=$thisproxy}
Set-ADUser -Identity $thissam -Add @{Proxyaddresses="SMTP:"+$thisnick}
Set-ADUser -Identity $thissam -Add @{Proxyaddresses="smtp:"+$thismail}
Set-ADUser -Identity $thissam -clear mail
Set-ADUser -Identity $thissam -Add @{mail="$thisnick"}


$updateuser = Get-ADuser $ff -Properties targetAddress, mail, sn, givenName, proxyaddresses | select-object targetAddress, mail, sn, givenName, @{n = "proxyAddress"; e = { $_.proxyAddresses | Where-object { $_ -clike "SMTP:*" } } } , mailNickname
$updatetime = get-date -Format MM-dd-yy-HH:ss
$updateproxy = $updateuser.proxyaddress
$updatemail = $updateuser.mail
$updatemsg = "$updatetime $thisupn now reads $updatemail | $updateproxy"
    Write-host $updatemsg
    $updatemsg | out-file log.txt -append


}
