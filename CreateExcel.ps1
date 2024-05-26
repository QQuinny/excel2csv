# example for data handling.
# Import the module
Import-Module ImportExcel

# Define the path for the Excel file
$excelFilePath = "C:\Users\jiaoxue\Documents\0526_24\Excel_twosheets.xlsx"

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
$data_1 | Export-Excel -Path $excelFilePath -WorksheetName "Sheet1"
$data_2 | Export-Excel -Path $excelFilePath -WorksheetName "Sheet2"

Write-Output "Example Excel file has been created: $excelFilePath"

Pause

