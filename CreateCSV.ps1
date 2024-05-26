# example for data handling.
# Import the module
# Import-Module ImportExcel

# Define the path for the Excel file
$csv_FilePath = "C:\Users\jiaoxue\Documents\0526_24\csv_file.csv"

# Create sample data
$data_1 = @(
    [PSCustomObject]@{ Name = "John Doe"; Age = 30; City = "New York" }
    [PSCustomObject]@{ Name = "Jane Smith"; Age = 25; City = "Los Angeles" }
    [PSCustomObject]@{ Name = "Sam Brown"; Age = 35; City = "Chicago" }
)

$data_2 = @(
    [PSCustomObject]@{ Name = "John Doe"; Age = 30; City = "New York" }
    [PSCustomObject]@{ Name = "Jane Smith"; Age = 25; City = "Los Angeles" }
    [PSCustomObject]@{ Name = "Sam Brown"; Age = 35; City = "Chicago" }
)

# Export data to Excel
$data_1 | Export-Csv  -Path $csv_FilePath -Append 
$data_2 | Export-Csv  -Path $csv_FilePath -Append

Write-Output "Example Excel file has been created: $csv_FilePath"

Pause
