# Import the ImportExcel module
Import-Module ImportExcel

# Import the Excel file
$excelFilePath = "C:\Users\jiaoxue\Documents\0526_24\Multi_Test_Files.xlsx"




$worksheetNames = @("sheet1", "sheet3")        # how to make array in powershell

write-Output $worksheetNames
# Loop through each worksheet and import data
 foreach ($worksheet in $worksheetNames) {
   $data = Import-Excel -Path $excelFilePath -WorksheetName $worksheet
   $selectedColumns = $data| Select-Object -Property Name, City

   # Output the selected columns
   $selectedColumns | Format-Table -AutoSize

   #$selectedColumns | Export-Csv -Path "C:\Users\jiaoxue\Documents\0526_24\Select_ColumnsfromExcel.csv" -NoTypeInformation -Append # Excel to Csv
   
   # Export the selected columns to a new Excel file
   $selectedColumns | Export-Excel -Path "C:\Users\jiaoxue\Documents\0526_24\Select_ColumnsfromExcel.xlsx" -WorksheetName "Park1" -Append  # Excel to Excel

}


pause

