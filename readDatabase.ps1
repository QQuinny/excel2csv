# Define the connection parameters
$serverName = "your_server_name"
$databaseName = "your_database_name"
$userName = "your_username" # Optional for Windows Authentication
$password = "your_password" # Optional for Windows Authentication

# Build the connection string
$connectionString = "Server=$serverName;Database=$databaseName;Integrated Security=True;" # For Windows Authentication
# $connectionString = "Server=$serverName;Database=$databaseName;User Id=$userName;Password=$password;" # For SQL Authentication

# Define the query
$query = "SELECT * FROM your_table_name"

# Create a new SQL connection
$connection = New-Object System.Data.SqlClient.SqlConnection
$connection.ConnectionString = $connectionString

# Open the connection
$connection.Open()

# Create a new SQL command
$command = $connection.CreateCommand()
$command.CommandText = $query

# Execute the command and read the data
$reader = $command.ExecuteReader()

# Create a DataTable to hold the results
$dataTable = New-Object System.Data.DataTable
$dataTable.Load($reader)

# Close the reader and the connection
$reader.Close()
$connection.Close()

# Output the results
$dataTable | Format-Table -AutoSize

# Optionally, you can export the data to a CSV file
$dataTable | Export-Csv -Path "C:\path\to\output\file.csv" -NoTypeInformation
