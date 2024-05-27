Install-Module -Name JoinObject -Scope CurrentUser

# Define the first dataset
$employees = @(
    [PSCustomObject]@{ EmployeeID = 1; Name = "Alice"; DepartmentID = 1 }
    [PSCustomObject]@{ EmployeeID = 2; Name = "Bob"; DepartmentID = 2 }
    [PSCustomObject]@{ EmployeeID = 3; Name = "Charlie"; DepartmentID = 1 }
)

# Define the second dataset
$departments = @(
    [PSCustomObject]@{ DepartmentID = 1; DepartmentName = "HR" }
    [PSCustomObject]@{ DepartmentID = 2; DepartmentName = "IT" }
    [PSCustomObject]@{ DepartmentID = 3; DepartmentName = "Finance" }
)
