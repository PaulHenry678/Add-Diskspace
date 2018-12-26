#$servers = import-csv -Path C:\temp\serverstoexpand.csv
$servers = (import-csv -Path C:\temp\serverstoexpand.csv) | select -ExpandProperty servername
$total = $servers | measure | select -expandproperty count
$i = 1

foreach ($server in $servers) {
Write-Host -ForegroundColor DarkYellow  "On $server, $i of $total"
Add-DiskSpace $server
$i++
}