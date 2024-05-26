# Import the module, this data is to use based on the "readAll.ps1"
Import-Module ImportExcel

# Define the path to the Excel file
$excelFilePath = "C:\Users\jiaoxue\Documents\excel2csv\example.xlsx"

# Check if the Excel file exists
if (-Not (Test-Path -Path $excelFilePath)) {
    Write-Error "The specified Excel file does not exist: $excelFilePath"
    exit 1
}

# Get the list of worksheet names
$worksheetNames = (Get-ExcelSheetInfo -Path $excelFilePath).Name

# Loop through each worksheet and import data
foreach ($worksheet in $worksheetNames) {
    Write-Output "Reading data from worksheet: $worksheet"
    $data = Import-Excel -Path $excelFilePath -WorksheetName $worksheet
    Write-Output "Data from {$worksheet}:"
    # Write-Output $data

    # You can process the data further as needed
    $data | Export-Csv -Path "C:\Users\jiaoxue\Documents\excel2csv\allsheets.csv" -NoTypeInformation -Append
}