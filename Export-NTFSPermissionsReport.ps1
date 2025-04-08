<#
.SYNOPSIS
    Exports NTFS folder permissions (ACLs) recursively to a CSV file.

.DESCRIPTION
    This script recursively scans a specified folder path and collects detailed NTFS permissions 
    (ACLs) on each folder. The results are saved to a timestamped CSV file with the computer name 
    and target folder encoded into the filename. Ideal for documenting and auditing file system permissions.

.NOTES
    Author: James Draper
    Last Modified: 2025-04-07

.EXAMPLE
    .\Export-NTFSPermissionsReport.ps1

.EXAMPLE
    .\Export-NTFSPermissionsReport.ps1 -FolderPath "F:\HomeDrive\$"

    Use backtick (`) to escape `$` in folder paths if needed, e.g., "F:\Profile`$"
#>

param (
    [Parameter(Mandatory = $false)]
    [string]$FolderPath = "E:\",  # Default path, can be overridden
    [string]$ReportPath = "C:\scripts" # Default output folder
)

# Ensure the folder exists
if (-not (Test-Path -Path $FolderPath)) {
    Write-Error "Folder path '$FolderPath' does not exist. Please check the path and try again."
    return
}

# Prepare an array list to store the permission details
$permissions = [System.Collections.Generic.List[PSObject]]::new()

# Function to process folders recursively
function Get-PermissionsRecursively {
    param (
        [string]$TargetPath
    )

    try {
        $acl = Get-Acl -Path $TargetPath

        $acl.Access | ForEach-Object {
            $permissions.Add([PSCustomObject]@{
                IdentityReference  = $_.IdentityReference
                FileSystemRights   = $_.FileSystemRights
                AccessControlType  = $_.AccessControlType
                IsInherited        = $_.IsInherited
                InheritanceFlags   = $_.InheritanceFlags
                PropagationFlags   = $_.PropagationFlags
                Path               = $TargetPath
            })
        }

        # Recurse into subdirectories
        Get-ChildItem -Path $TargetPath -Directory -Force -ErrorAction SilentlyContinue | ForEach-Object {
            Get-PermissionsRecursively -TargetPath $_.FullName
        }
    } catch {
        Write-Warning "Failed to retrieve permissions for '$TargetPath': $_"
    }
}

# Start recursion
Get-PermissionsRecursively -TargetPath $FolderPath

# Sanitize the folder path for use in filenames
$sanitizedPath = ($FolderPath -replace '[\\:\*?"<>|]', '_') -replace '__+', '_'
$timestamp = Get-Date -Format "yyyyMMdd"
$serverName = $env:COMPUTERNAME
$csvFileName = "PermissionsReport_${serverName}_${sanitizedPath}_${timestamp}.csv"
$csvFilePath = Join-Path -Path $ReportPath -ChildPath $csvFileName

# Export the data
$permissions | Export-Csv -Path $csvFilePath -NoTypeInformation -Encoding UTF8

Write-Output "`n✅ Permissions exported successfully to:`n$csvFilePath"
