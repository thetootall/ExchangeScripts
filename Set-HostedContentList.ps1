Param(
   [Parameter(Mandatory=$True,Position=1)]
   [string]$ruleName,
  
   [Parameter(Mandatory=$True)]
   [string]$ListFilePath
)

#Update this policy to account for emails as follows:

# Get-HostedContentFilterPolicy -AllowedSenders
# Get-HostedConnectionFilterPolicy -IPAllowList

#Read the contents of the text file into an array

Clear-Variable completeList -ErrorAction SilentlyContinue
Clear-Variable safeDomainList -ErrorAction SilentlyContinue
Clear-Variable newSafeDomainList -ErrorAction SilentlyContinue

[array]$newsafeDomainList = Get-Content $ListFilePath
#If the rule already exists update the existing allowed sender domains, else create a new rule.

if (Get-HostedContentFilterPolicy $ruleName -EA SilentlyContinue)
{
  "Updating existing rule..."
  [array]$safeDomainList = Get-HostedContentFilterPolicy $ruleName |select -ExpandProperty AllowedSenderDomains
  [array]$completeList = $safeDomainList + $newSafeDomainList
  [array]$completeList = $completeList | select -uniq | sort

  Set-HostedContentFilterPolicy $ruleName -AllowedSenders $null

  write-host $completeList
  $ready = read-host "Press Y to commit"
  If ($ready -eq "Y"){
  
  $completeList | foreach {Set-HostedContentFilterPolicy $ruleName -AllowedSenderDomains @{add=$_}} 
  }
}
else
{
    Write-host "No Content Filter Policy found; please review via Exchange Admin Center"
}

Set-HostedContentFilterPolicy $ruleName | fl AllowedSenderDomain
