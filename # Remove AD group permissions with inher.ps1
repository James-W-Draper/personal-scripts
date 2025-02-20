# PowerShell Script to Remove AD Group Permissions with Inheritance Check
# Version 3.0, Created 11-09-2024
#
# This script removes Read and Execute permissions for a specified AD group from a given root directory and its subdirectories,
# ensuring inherited permissions are considered.
#
# Prerequisites:
# - The executing user must have permission to modify ACLs.
# - Ensure PowerShell is running with administrative privileges.

# Define variables for permissions and folder path
$rootFolder = "J:\F_Drive\Users$"  # Replace with the top-level folder path
$adGroup = "cwglobal\SD_Admins_L2"  # Replace with your AD group name
$scriptPath = "C:\scripts"  # Define the directory for storing logs and reports

# Define variables for error logging
$date = Get-Date -Format "ddMMyyyy"
$datetime = Get-Date -Format "ddMMyyyy HH:mm"
$serverName = $env:COMPUTERNAME

# Sanitize folder name for use in log filenames
$sanitizedRootFolder = $rootFolder -replace '[\\:]', '_' -replace '__+', '_'

# Define log file paths
$errorLogFile = "${scriptPath}\Errors_${serverName}_${sanitizedRootFolder}_${date}.txt"
$transcriptFile = "${scriptPath}\Transcript_${serverName}_${sanitizedRootFolder}_${date}.txt"

# Output constructed file paths for verification
Write-Output "Logging errors to: $errorLogFile"
Write-Output "Logging transcript to: $transcriptFile"

# Start transcript to capture script execution details
Start-Transcript -Path $transcriptFile

# Log key variables
Write-Host "Script started at ${datetime} with the following parameters:" -ForegroundColor Cyan
Write-Host "  Root Folder: $rootFolder"
Write-Host "  AD Group: $adGroup"
Write-Host "  Script Path: $scriptPath"
Write-Host "  Error Log: $errorLogFile"

# Function to remove Read and Execute access from an AD group
function Remove-ReadAccessFromAdGroup {
    param (
        [string]$path,
        [string]$adGroup
    )

    # Retrieve ACL information for the target path
    $acl = Get-Acl -Path $path
    $existingRule = $acl.Access | Where-Object {
        $_.IdentityReference -eq $adGroup -and 
        $_.FileSystemRights -eq [System.Security.AccessControl.FileSystemRights]::ReadAndExecute -and 
        $_.IsInherited -eq $true
    }

    if ($existingRule) {
        Write-Output "Removing permissions for $adGroup from $path (Inherited)"
        $acl.RemoveAccessRule($existingRule)
        Set-Acl -Path $path -AclObject $acl
    } else {
        Write-Output "No inherited permissions found for $adGroup on $path. Skipping..."
    }

    # Process child directories recursively
    Get-ChildItem -Path $path -Recurse -Directory | ForEach-Object {
        try {
            if ($null -ne $_) {
                $childAcl = Get-Acl $_.FullName
                $childExistingRule = $childAcl.Access | Where-Object {
                    $_.IdentityReference -eq $adGroup -and 
                    $_.FileSystemRights -eq [System.Security.AccessControl.FileSystemRights]::ReadAndExecute -and 
                    $_.IsInherited -eq $true
                }

                if ($childExistingRule) {
                    Write-Output "Removing permissions for $adGroup from $($_.FullName) (Inherited)"
                    $childAcl.RemoveAccessRule($childExistingRule)
                    Set-Acl -Path $_.FullName -AclObject $childAcl
                } else {
                    Write-Output "No inherited permissions found for $adGroup on $($_.FullName). Skipping..."
                }
            }
        } catch {
            # Log errors encountered during processing
            $errorMsg = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Error processing $_.FullName: $($_.Exception.Message)"
            $errorMsg | Out-File -Append -FilePath $errorLogFile
        }
    }
}

# Execute the function to remove permissions
Write-Output "Initiating permission removal process..."
Remove-ReadAccessFromAdGroup -path $rootFolder -adGroup $adGroup

# End transcript
Stop-Transcript
Write-Host "Script execution completed." -ForegroundColor Green
