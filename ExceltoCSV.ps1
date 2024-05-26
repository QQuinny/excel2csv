# created by Park // select one sheet in a file, make it based on the Original document of "ConvertExcel2Csv.ps1" , can be generated to change of $worksheetName = "sheet1 or sheet2"

    $excelFilePath = "C:\Users\jiaoxue\Documents\0526_24\Excel_twosheets.xlsx"
    $csvOutputPath = "C:\Users\jiaoxue\Documents\0526_24\Covert_Excelinto_CSV.csv"
    $worksheetName = "sheet3"

     

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
      # $data = Import-Excel -Path $excelFilePath -WorksheetName "sheet3"
    } else {
        $data = Import-Excel -Path $excelFilePath 
    }
} catch {
    Write-Error "Failed to import Excel file: $_"
    exit 1
}

# Export the data to CSV
try {
    $data | Export-Csv -Path $csvOutputPath -NoTypeInformation    # | pipeline
    Write-Output "Excel file has been successfully converted to CSV: $csvOutputPath"
} catch {
    Write-Error "Failed to export data to CSV: $_"
    exit 1
}

# "ctl + /" to make all line of comments
# Why Use -NoTypeInformation?
# Compatibility: Some applications that consume CSV files may not handle the type information line correctly. Using -NoTypeInformation ensures the CSV file is in a more standard format.
# Readability: Removing the type information makes the CSV file easier to read for users who open it in text editors or spreadsheet applications like Microsoft Excel.
# Storage and Scripting: If you are processing the CSV file in other scripts or tools that don't need the type information, omitting it can simplify the file structure.
