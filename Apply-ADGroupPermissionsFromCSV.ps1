# Path to the CSV file containing folder paths
$csvPath = "C:\Scripts\Project Folder Owners - Shortened - WIP- ukgs2file1.csv"

# Path where the permission validation results will be exported as a CSV
$outputCsvPath = "C:\Scripts\permission_validation.csv"

# Define the AD group for which read-only permissions will be applied (update with your AD group name)
$adGroup = "CWGLOBAL\SD_Admins_L2"

# Import the CSV file into a variable. Each row from the CSV will represent a folder path.
$foldersList = Import-Csv $csvPath

# Initialize an empty array to store results for export
$results = @()

# Loop through each row (folder entry) in the CSV file
foreach ($entry in $foldersList) {
    # Extract the folder path from the 'Consolidated Server path' column of the CSV
    $folderPath = $entry.'Consolidated Server path'

    # Check if the folder exists before proceeding
    if (-not (Test-Path $folderPath)) {
        Write-Output "Folder does not exist: $folderPath"
        $results += [pscustomobject]@{
            'Folder Path'     = $folderPath
            'AD Group'        = $adGroup
            'Permission Set'  = "Failed"
            'Error Message'   = "Folder does not exist"
        }
        continue
    }

    try {
        # Retrieve the Access Control List (ACL) for the current folder
        $acl = Get-Acl $folderPath

        # Create a new permission rule for the specified AD group with read-only permissions
        $permission = New-Object System.Security.AccessControl.FileSystemAccessRule(
            $adGroup, 
            "ReadAndExecute", 
            "ContainerInherit,ObjectInherit",  # Inheritance flags for the group
            "None",                            # Disable propagation of inheritance
            "Allow"                            # This is an allow rule (not deny)
        )

        # Add the newly created permission rule to the ACL of the folder
        $acl.AddAccessRule($permission)

        # Apply the modified ACL back to the folder (this sets the new permission on the folder)
        Set-Acl -Path $folderPath -AclObject $acl

        # Re-fetch the ACL to confirm that the permission was applied correctly
        $acl = Get-Acl $folderPath

        # Initialize a flag to check if the permission was applied successfully
        $permissionApplied = $false

        # Loop through each access rule in the ACL to see if our rule for the AD group was successfully added
        foreach ($accessRule in $acl.Access) {
            # Use -match to check for a match of "ReadAndExecute" in case it is combined with other rights
            if ($accessRule.IdentityReference -eq $adGroup -and ($accessRule.FileSystemRights -band [System.Security.AccessControl.FileSystemRights]::ReadAndExecute) -ne 0) {
                $permissionApplied = $true
                break  # Exit the loop if the permission is found
            }
        }

        # Record the result in the array. If the permission was applied, log "Success", otherwise "Failed"
        $results += [pscustomobject]@{
            'Folder Path'     = $folderPath
            'AD Group'        = $adGroup
            'Permission Set'  = if ($permissionApplied) { "Success" } else { "Failed" }
        }

        # Output the result to the console for real-time feedback
        if ($permissionApplied) {
            Write-Output "Permissions successfully applied to folder: $folderPath for AD Group: $adGroup"
        } else {
            Write-Output "Failed to apply permissions to folder: $folderPath for AD Group: $adGroup"
        }

    } catch {
        # If an error occurs, capture the detailed error message and log it
        $errorMessage = $_.Exception.Message  # Get the exception message
        $errorStackTrace = $_.Exception.StackTrace  # Get the full stack trace for debugging (optional)

        # Record the error in the result array, including the error message and stack trace
        $results += [pscustomobject]@{
            'Folder Path'     = $folderPath
            'AD Group'        = $adGroup
            'Permission Set'  = "Error"
            'Error Message'   = $errorMessage
            'Stack Trace'     = $errorStackTrace  # Optional, can be removed if you don't need the stack trace
        }

        # Output the detailed error message to the console for real-time feedback
        Write-Output "Error processing folder: $folderPath - $errorMessage"
    }
}

# Once the loop is finished, export the results to a CSV file at the specified path
$results | Export-Csv -Path $outputCsvPath -NoTypeInformation

# Inform the user that the process is complete and where the results are stored
Write-Output "Permission validation results exported to $outputCsvPath"
