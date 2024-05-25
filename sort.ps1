# Define the connection parameters
$serverName = "your_server_name"
$databaseName = "your_database_name"
$userName = "your_username" # Optional for Windows Authentication
$password = "your_password" # Optional for Windows Authentication

# Build the connection string
$connectionString = "Server=$serverName;Database=$databaseName;Integrated Security=True;" # For Windows Authentication
# $connectionString = "Server=$serverName;Database=$databaseName;User Id=$userName;Password=$password;" # For SQL Authentication

# Define the query with ORDER BY
$query = "SELECT City, Name, Age  FROM People ORDER BY Age ASC, Name DESC"

# Create a new SQL connection
$connection = New-Object System.Data.SqlClient.SqlConnection
$connection.ConnectionString = $connectionString

try {
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

    # Output the results
    $dataTable | Format-Table -AutoSize

    # Optionally, you can export the data to a CSV file
    $dataTable | Export-Csv -Path "C:\path\to\output\sorted_data.csv" -NoTypeInformation
} catch {
    Write-Error "An error occurred: $_"
} finally {
    if ($reader) { $reader.Close() }
    if ($connection) { $connection.Close() }
}
