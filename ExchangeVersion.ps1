#Get Exchange Server Versions
#Be sure Powershell Remoting is enabled
#To enable on each server run winrm quickconfig
#Created by Chris Blackburn http://memphistech.net

$exchangeservers = Get-ExchangeServer
 
$report = @()
 
foreach ($srv in $exchangeservers)
{

$srv = Get-ExchangeServer $srv
 
$server = $srv.Name
if ($srv.AdminDisplayVersion -match "Version 14") {$ver = "V14"}
if ($srv.AdminDisplayVersion -match "Version 15") {$ver = "V15"}
 
    Write-Host "Checking $server"
 
    $installpath = $null
 
    try
    {
        $installpath = Invoke-Command –Computername $server -ScriptBlock {$env:ExchangeInstallPath} -ErrorAction STOP
    }
    catch
    {
        Write-Warning $_.Exception.Message
        $installpath = "Unable to connect to server"
    }

Write-host "Install Path is: " + $installpath
$Path = $installpath + "Bin\ExSetup.exe"
$fileversion = (Get-Command $Path).FileVersionInfo
Write-Host "File Version is: " + $FileVersion.fileversion

    $serverObj = New-Object PSObject
	$serverObj | Add-Member NoteProperty -Name "Server Name" -Value $server
	$serverObj | Add-Member NoteProperty -Name "Install Path" -Value $installpath
	$serverObj | Add-Member NoteProperty -Name "Server Role" -Value $srv.serverrole
	$serverObj | Add-Member NoteProperty -Name "Server Version" -Value $fileversion.fileversion
 
   
    $report += $serverObj  

Clear-variable srv
Clear-variable server
Clear-variable installpath
Clear-variable ver
}
 
$report | export-csv exchangeversions.csv -notype