Param(
   [Parameter(Mandatory=$True,Position=1)]
   [string]$ruleName,
  
   [Parameter(Mandatory=$True)]
   [string]$domainListFilePath
)
#Read the contents of the text file into an array
[Array]$newsafeDomainList = Get-Content $domainListFilePath
#If the rule already exists update the existing allowed sender domains, else create a new rule.
if (Get-TransportRule $ruleName -EA SilentlyContinue)
{
  "Updating existing rule..."
  [array]$safeDomainList = Get-TransportRule $ruleName |select -ExpandProperty ExceptIfRecipientAddressContainsWords
  [array]$completeList = $safeDomainList + $newSafeDomainList
  [array]$completeList = $completeList | select -uniq | sort
  write-host $completelist
  $ready = read-host "Press Y to commit"
  If ($ready -eq "Y"){
  set-TransportRule $ruleName -ExceptIfRecipientAddressContainsWords $completeList
}
