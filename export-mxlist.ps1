# A customer wanted to get an output of MX records for the domain names in their CRM contacts to see who the mail providers were
# For this to work properly, be sure you populate a CSV with a header called "domain" and each row with a valid domain name.
$domains = import-csv "cust-domains.csv"
foreach ($item in $domains){
$domain = $item.domain
Write-host $domain
$domain | resolve-dnsname -Type MX -Server 8.8.8.8 | where {$_.QueryType -eq "MX"} | Select Name,NameExchange | export-csv cust-output.csv -notype -append
}
