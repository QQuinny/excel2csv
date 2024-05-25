$csvData = Import-Csv -Path .\example.csv


# foreach ($row in $csvData) { Write-Output $row.name  $row.Age $row.City }
$csvData | Format-Table -AutoSize

Pause