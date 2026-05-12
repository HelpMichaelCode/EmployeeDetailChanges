# Importing Employee Details
$employeeDetails = Import-Csv -Path "$env:OneDrive\Documents\PowerShell_Projects\OriginalPeopleChanges.csv"

# Cleaning the data
$employeeDetails = $employeeDetails |
    ForEach-Object {
        $_.NewName     = $_.NewName -replace '^(?i)(mr|mrs)\s+', ''
        $_.NewManager  = $_.NewManager -replace '^(?i)(mr|mrs)\s+', ''
        $_.OriginalName     = $_.OriginalName -replace '^(?i)(mr|mrs)\s+', ''
        $_.OriginalManager  = $_.OriginalManager -replace '^(?i)(mr|mrs)\s+', ''

        $NewEmpID = $_.NewUserEmail.Split('@')[0].ToLower()
        $_ | Add-Member -MemberType NoteProperty -Name NewEmpID -Value $NewEmpID -Force
        
        $_  
    }

# Created a hash table for old Employee details.
# This allows us to access the data directly, instead of looping through the entire .csv file

$empLookup = @{}
foreach ($emp in $employeeDetails) {
    
    # For each employee, create an object and add it into the hash table
    $empLookup[$emp.NewEmpID] = $emp
}

# Loop through the list of employees
$updatedEmpDetails = $employeeDetails | ForEach-Object {
    
    # Grab the employee details and put into an object to access within the IF statements
    $employee = $empLookup[$_.NewEmpID]


    # Check if the employee exists
if ($employee) {
        
        # These IF statements are essentially replacing the Original values with the New values from the .CSV file

        if ($employee.OriginalName -eq "" -or
            $employee.OriginalName -ne $employee.NewName) {

            $employee.OriginalName = $employee.NewName
        }

        if ($employee.OriginalPosition -eq "" -or
            $employee.OriginalPosition -ne $employee.NewPosition) {

            $employee.OriginalPosition = $employee.NewPosition
        }

        if ($employee.OriginalDepartment -eq "" -or
            $employee.OriginalDepartment -ne $employee.NewDepartment) {

            $employee.OriginalDepartment = $employee.NewDepartment
        }

        if ($employee.OriginalManager -eq "" -or
            $employee.OriginalManager -ne $employee.NewManager) {

            $employee.OriginalManager = $employee.NewManager
        }


        if (($employee.OriginalUserEmail -eq "" -or
            $employee.OriginalUserEmail -ne $employee.NewUserEmail) -and
            $employee.NewUserEmail -ieq "User Email Unknown") {

            Write-Output "$($employee.NewName)'s new email is unknown. Email Address won't be changed."
             
        } elseif ($employee.OriginalUserEmail -eq "" -or
            $employee.OriginalUserEmail -ne $employee.NewUserEmail) {
            
                $employee.OriginalUserEmail = $employee.NewUserEmail

            }

        if ($employee.NewPosition -eq "Leaver") {
        
            $employee.OriginalName = "Disabled"
            $employee.OriginalPosition = "Disabled"
            $employee.OriginalDepartment = "Disabled"
            $employee.OriginalManager = "Disabled"
            $employee.OriginalUserEmail = "Disabled"

        }
    }

    # If the employee does not exist, print this out
 else {
        
        Write-Output "$($employee.NewName) does not exist."
    
    }
}

# Export the updated Employee Details to a brand new .csv file
$employeeDetails | 
    Select-Object NewName, NewPosition, NewDepartment, NewManager, NewUserEmail, 
                  OriginalName, OriginalPosition, OriginalDepartment, OriginalManager, OriginalUserEmail |
    Export-Csv -Path "$env:OneDrive\Documents\PowerShell_Projects\UpdatedPeopleChanges.csv" -NoTypeInformation