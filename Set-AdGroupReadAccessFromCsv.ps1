<#
.SYNOPSIS
Grants read-only NTFS permissions to a specified AD group across multiple folders listed in a CSV.

.DESCRIPTION
This script reads a list of folder paths from a CSV file and applies "Read and Execute" NTFS permissions
to a specified Active Directory group on each folder. It verifies and logs success or failure for each path.

.PARAMETER csvPath
The path to the input CSV file. The CSV must contain a column named 'Consolidated Server path'.

.PARAMETER outputCsvPath
The path to save the output report detailing success/failure and any errors.

.PARAMETER adGroup
The AD group to which permissions will be granted (DOMAIN\GroupName format).

.NOTES
- Requires administrative privileges
- Requires the folder paths to be accessible and writable
#>

# === CONFIGURATION ===
$csvPath = "C:\Scripts\ProjectFolders.csv"                     # <-- Path to the input CSV file
$outputCsvPath = "C:\Scripts\PermissionValidationReport.csv"   # <-- Path for export results
$adGroup = "DOMAIN\GroupName"                                  # <-- Replace with the AD group

# === BEGIN PROCESSING ===
$foldersList = Import-Csv $csvPath
$results = @()

foreach ($entry in $foldersList) {
    $folderPath = $entry.'Consolidated Server path'

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
        $acl = Get-Acl $folderPath

        $permission = New-Object System.Security.AccessControl.FileSystemAccessRule(
            $adGroup,
            "ReadAndExecute",
            "ContainerInherit,ObjectInherit",
            "None",
            "Allow"
        )

        $acl.AddAccessRule($permission)
        Set-Acl -Path $folderPath -AclObject $acl

        # Confirm permission was applied
        $acl = Get-Acl $folderPath
        $permissionApplied = $false

        foreach ($accessRule in $acl.Access) {
            if (
                $accessRule.IdentityReference -eq $adGroup -and
                ($accessRule.FileSystemRights -band [System.Security.AccessControl.FileSystemRights]::ReadAndExecute) -ne 0
            ) {
                $permissionApplied = $true
                break
            }
        }

        $results += [pscustomobject]@{
            'Folder Path'     = $folderPath
            'AD Group'        = $adGroup
            'Permission Set'  = if ($permissionApplied) { "Success" } else { "Failed" }
        }

        Write-Output "[$(Get-Date -Format 'HH:mm:ss')] $($permissionApplied ? '✔️' : '❌') $folderPath"

    } catch {
        $results += [pscustomobject]@{
            'Folder Path'     = $folderPath
            'AD Group'        = $adGroup
            'Permission Set'  = "Error"
            'Error Message'   = $_.Exception.Message
            'Stack Trace'     = $_.Exception.StackTrace
        }

        Write-Warning "Error processing folder: $folderPath - $($_.Exception.Message)"
    }
}

# Export results to CSV
$results | Export-Csv -Path $outputCsvPath -NoTypeInformation
Write-Output "`n✅ Permission validation results exported to:`n$outputCsvPath"
