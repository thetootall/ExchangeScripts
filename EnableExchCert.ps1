#Automatically enables Exchange Certificate based on subject name
#For updates and issues: https://github.com/thetootall/ExchangeScripts/blob/master/EnableExchCert.ps1
$varurl = Read-Host "Enter Certificate Subject"
$varsub = "CN=" + $varurl
$varthumb = Get-ChildItem -Recurse Cert:\LocalMachine\My | Where { $_.subject -eq $varsub -and $_.NotAfter -gt (Get-Date).AddDays(0) } | Select Thumbprint
$thumbprint = $varthumb.thumbprint
Enable-ExchangeCertificate -Thumbprint $thumbprint -Services POP,IMAP,IIS,SMTP
