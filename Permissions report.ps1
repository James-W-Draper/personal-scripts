# Permissions report
# Run this script on the local server
# Define the path to the folder
$folderPath = "E:\"
# If using a folderpath containing a $ symbol, use something like
# $folderPath = "F:\Drive_E\Profile`$"

# Prepare an empty array list to store the permission details
$permissions = [System.Collections.Generic.List[PSObject]]::new()

# Function to process folder and its subfolders
function Deploy-Permissions {
    param (
        [string]$folderPath
    )

    # Get ACL for the folder
    $acl = Get-Acl -Path $folderPath

    # Loop through each AccessRule in the ACL using ForEach-Object for pipeline processing
    $acl.Access | ForEach-Object {
        $permission = [PSCustomObject]@{
            IdentityReference  = $_.IdentityReference
            FileSystemRights   = $_.FileSystemRights
            AccessControlType  = $_.AccessControlType
            IsInherited        = $_.IsInherited
            InheritanceFlags   = $_.InheritanceFlags
            PropagationFlags   = $_.PropagationFlags
            Path               = $folderPath
        }
        $permissions.Add($permission)
    }

    # Get subfolders and process each one
    $subfolders = Get-ChildItem -Path $folderPath -Directory
    foreach ($subfolder in $subfolders) {
        Deploy-Permissions -folderPath $subfolder.FullName
    }
}

# Start processing from the root folder
Deploy-Permissions -folderPath $folderPath

# Sanitize the root path for use in file names
$sanitizedRootPath = $folderPath -replace '[\\:\*?"<>|]', '_' -replace '__+', '_'

# Get the current server name
$serverName = $env:COMPUTERNAME

# Get the current date in the format YYYYMMDD
$date = Get-Date -Format "yyyyMMdd"

# Combine to create the CSV file name
$csvFileName = "PermissionsReport_${serverName}_${sanitizedRootPath}_${date}.csv"

# Report Folder path
$Reportpath = "C:\scripts"

# Define the full path to save the CSV
$csvFilePath = "${Reportpath}\${csvFileName}"

# Export the permissions array to a CSV file
$permissions | Export-Csv -Path $csvFilePath -NoTypeInformation

Write-Output "Permissions exported to $csvFilePath"
