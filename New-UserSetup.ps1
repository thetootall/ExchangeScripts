Import-module ActiveDirectory

$firstname = Read-Host "Enter first name"
$lastname = Read-Host "Enter last name"
$usershort = Read-Host "Enter username"

#Customizable Variables
$domain = "@idri.org"
$vanitydomain = "@idriorg.mail.onmicrosoft.com"
$pathOU = "OU=IDRIUsers,OU=MyBusiness,OU=Users,DC=idri,DC=org"

#DO NOT MODIFY PAST THIS LINE!!!!

$whole = $firstname + "." + $lastname
$dispname = $firstname + " " + $lastname

$lastshort = $lastname.SubString(0,1)

#$usersam = $Firstname + $lastshort

#This builds the UPN
$usere = $whole + $Domain

#This builds the TargetAddress
$targetA = $usershort + $vanitydomain
#This builds the alias
$targetB = $usershort + $domain

Write-host "User first and Last name: $dispname"
Write-host "User email: $usere"
Write-host "User UPN: $targetB"
write-host "User SAM: $usershort"

$confirm = Read-host "Please enter Y if user information looks correct"

If ($confirm -eq "Y"){
#Creates the AD User
New-ADUser -GivenName $firstname -Surname $lastname -DisplayName $dispname -Name $dispname -UserPrincipalName $targetB -SamAccountName $usershort -path $pathOU -AccountPassword (ConvertTo-SecureString "abc123" -AsPlainText -Force) -Enabled $True

Write-host "Waiting 10 seconds" -Foregroundcolor Yellow
Sleep 10

#Look up newly created user
$adobject = Get-ADUser -Filter {UserPrincipalName -eq $targetB}

$adobject | Set-ADUser -Add @{proxyAddresses = "SMTP:" + $usere}
$adobject | Set-ADUser -Add @{proxyAddresses = "smtp:" + $targetA}
$adobject | Set-ADUser -Add @{proxyAddresses = "smtp:" + $targetB}

$adobject | Set-ADUser -Email $usere
$adobject | Set-ADUser -Add @{targetAddress = "SMTP:" + $targetA}

$adobject | Set-ADUser -Add @{msExchRecipientDisplayType = "-2147483642"}
$adobject | Set-ADUser -Add @{msExchRecipientTypeDetails = "2147483648"}
$adobject | Set-ADUser -Add @{msExchRemoteRecipientType = "3"}

#Perform ADSYNC on Server
Import-Module ADSYNC
Start-ADSyncSyncCycle -PolicyType Delta
}