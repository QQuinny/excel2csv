# created by Park
param (
    [string]$excelFilePath,
    [string]$csvOutputPath,
    [string]$worksheetName = $null
)

# Import the module
Import-Module ImportExcel

# Check if the Excel file exists
if (-Not (Test-Path -Path $excelFilePath)) {
    Write-Error "The specified Excel file does not exist: $excelFilePath"
    exit 1
}

# Import the Excel file
try {
    if ($worksheetName) {
        $data = Import-Excel -Path $excelFilePath -WorksheetName $worksheetName
    } else {
        $data = Import-Excel -Path $excelFilePath 
    }
} catch {
    Write-Error "Failed to import Excel file: $_"
    exit 1
}

# Export the data to CSV
try {
    $data | Export-Csv -Path $csvOutputPath -NoTypeInformation
    Write-Output "Excel file has been successfully converted to CSV: $csvOutputPath"
} catch {
    Write-Error "Failed to export data to CSV: $_"
    exit 1
}

pause

#Test it in singl file.
