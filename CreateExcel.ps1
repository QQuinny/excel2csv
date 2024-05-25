# Import the module
Import-Module ImportExcel

# Define the path for the Excel file
$excelFilePath = "C:\Users\jiaoxue\Documents\excel2csv\example.xlsx"

# Create sample data
$data = @(
    [PSCustomObject]@{ Name = "John Doe"; Age = 30; City = "New York" }
    [PSCustomObject]@{ Name = "Jane Smith"; Age = 25; City = "Los Angeles" }
    [PSCustomObject]@{ Name = "Sam Brown"; Age = 35; City = "Chicago" }
)

# Export data to Excel
$data | Export-Excel -Path $excelFilePath -WorksheetName "Sheet1"

Write-Output "Example Excel file has been created: $excelFilePath"

Pause
