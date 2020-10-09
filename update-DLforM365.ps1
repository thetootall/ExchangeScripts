#CSV will be formatted with two column headers: "email" with addresses of Shared mailboxes, and "group" with the name of the group
#$file = Read-host "type file path (if not in current folder)"
$csvlist = import-csv "grouplist.csv"
ForEach ($item in $csvlist) {
#start L1
$g = $item.group
write-host "Adding permission for $g"
Enable-DistributionGroup $g
Set-ADGroup -Identity $g -GroupScope Universal
#end L1
}
