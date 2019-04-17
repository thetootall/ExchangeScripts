#https://community.spiceworks.com/topic/1940442-add-bulk-proxyaddresses-attribute

#CSV MAP
#SamAccountName
#cblackb
Import-module ActiveDirectory

#---------------Set our Runtime Variables
#dumpuser compares mail attribute to primary PROXY address
$dumpuser = "NO"

#dumpuseroutput logs the output of the dumpuser process
$dumpuseroutput = "SMTPoutput.txt"

#updatedisplay triggers the preview of the SMTP update process
$updatedisplay = "YES"

#updateExecute triggers the AD update
$updateExecute = "YES"

#---------------End the Runtime variables
If ($dumpuser -eq "YES"){$users = Get-ADUser -SearchBase "OU=Users,OU=ENT,DC=domain,DC=local" -Filter {mail -ne $null}}
If ($dumpuser -eq "NO"){$users = Import-Csv "SMTPaddlUsers.csv"}

ForEach ($Item in $users){
#$thisuser = $item.("username")
$thisuser = $item.("userprincipalname")
#Write-host "Processing $thisuser" -ForegroundColor Green

clear-variable myuser*
$myuser = Get-AdUser -filter {userprincipalname -eq $thisuser} -properties givenname,sn,displayname,mail,proxyaddresses

$myuserfirst = $myuser.givenName
$myuserlast = $myuser.SN
$myuserdisp = $myuser.displayname
$myusersam = $myuser.SamAccountName
$myusermail = $myuser.mail
$myuserproxy = $myuser.proxyaddresses
$myuserprimsmtp = $myuserproxy -clike "*SMTP*"


If ($dumpuser -eq "YES"){
$dumparr = "$myusermail" + "," + "$myuserprimsmtp"
Write-host $dumparr
$dumparr | out-file $dumpuseroutput -Append}
#End output SMTP list

#start the display of the update process
If ($updatedisplay -eq "YES"){

#define domain here
$domain="@domain.com"
$primail = $myusermail.split("@")[0]
WRite-host "Address split = $primail" -ForegroundColor Green

#building proxies
$PriSMTP="SMTP:"
$AltSMTP="smtp:"
$OnMS = "@tenant.onmicrosoft.com"
$OnMS2 = "@tenant.mail.onmicrosoft.com"
$NameDotNo = $myuserfirst + $myuserlast
$NameDotYes = $myuserfirst + "." + $myuserlast
$NameDotYes = $NameDotYes.Replace(' ','')
$NameDotNoMail = $NameDotNo + $domain
$NameDotNoMail = $NameDotNoMail.Replace(' ','')

#adding all
$MAIL=$primail + $domain
$OnMS=$AltSMTP + $myusersam + $OnMS
$SMTP1=$PriSMTP + $MAIL
$SMTP2=$AltSMTP + $NameDotYes + $domain
$SMTP3=$AltSMTP + $myusersam + $OnMS2
$SMTP4=$AltSMTP + $myusersam + $domain
$SIPAddress = "SIP:" + $MAIL

#Start the process of displaying SMTP addresses
Write-host "Processing $myuserdisp" -ForegroundColor Yellow

Write-host $SMTP1
Write-host $smtp2
Write-host $smtp3
Write-host $smtp4
Write-host $onMS
Write-host $SIPAddress

If ($mail -ne $NameDotNoMail){
Write-host "Given vs Preferred Name Mismatch!!!" -ForegroundColor Red
$NameDotNoMail=$AltSMTP+$NameDotNoMail
Write-host "Adding additional alias $NameDotNoMail" -ForegroundColor Cyan
#End name alias
}

Write-host "-----------------------------------------" -ForegroundColor Cyan

#Start the live update of AD with new variables
If ($updateexecute -eq "YES"){
Get-ADUser $myusersam | set-aduser -Clear ProxyAddresses
Get-ADUser $myusersam | set-aduser -Add @{Proxyaddresses=$SMTP1}
Get-ADUser $myusersam | set-aduser -Add @{Proxyaddresses=$smtp2}
Get-ADUser $myusersam | set-aduser -Add @{Proxyaddresses=$smtp3}
Get-ADUser $myusersam | set-aduser -Add @{Proxyaddresses=$smtp4}
Get-ADUser $myusersam | set-aduser -Add @{Proxyaddresses=$onMS}

#check for given vs preferred name
If ($mail -ne $NameDotNoMail){
Get-ADUser $myusersam | set-aduser -Add @{Proxyaddresses=$NameDotNoMail}
#End name alias
}

Get-ADUser $myusersam | set-aduser -Replace @{'msRTCSIP-PrimaryUserAddress'=$SIPAddress}

#Get-ADUser $myusersam | set-aduser -Replace @{userprincipalname=$mail}


#end updateexecuite loop
}
#end updatedisplay loop
}
#end for loop
}
