<#
.SYNOPSIS
Grants "Read and Execute" NTFS permissions to a specified Active Directory group across a folder structure.

.DESCRIPTION
This script applies inherited Read & Execute permissions for an AD group on a specified root folder
and all subdirectories. It checks whether the group already has inherited access to avoid duplication.

Useful for onboarding support teams or access delegation scenarios.

The script:
- Adds access rules only if the group doesn't already have inherited rights
- Recursively processes all directories under the given path
- Logs progress to the console
- Records a transcript and error log for review

.PARAMETER rootFolder
Root folder where permission application begins (recurse through all subfolders)

.PARAMETER adGroup
The AD group (in DOMAIN\GroupName format) to receive permissions

.PARAMETER scriptPath
Directory to store the transcript and error logs

.NOTES
- Must be run with administrative privileges
- Script version: 3.0 (generalised)
#>

# === CONFIGURATION ===
$rootFolder = "D:\Path\To\RootFolder"        # <-- Update this to your target folder
$adGroup = "domain\GroupName"                # <-- Update this to your AD group
$scriptPath = "C:\Scripts"                   # <-- Folder for logs

# === RUNTIME METADATA ===
$date = Get-Date -Format "yyyyMMdd"
$datetime = Get-Date -Format "yyyy-MM-dd HH:mm"
$serverName = $env:COMPUTERNAME

# Sanitize folder name for safe use in filenames
$folderLabel = ($rootFolder -replace '[\\:\$]', '_') -replace '__+', '_'

# Log file paths
$errorLogFile = Join-Path $scriptPath "Errors_${serverName}_${folderLabel}_${date}.txt"
$transcriptFile = Join-Path $scriptPath "Transcript_${serverName}_${folderLabel}_${date}.txt"

# Output file paths
Write-Output "Logging transcript to: $transcriptFile"
Write-Output "Logging errors to: $errorLogFile"

# Start transcript
Start-Transcript -Path $transcriptFile

Write-Host "Script started at ${datetime}" -ForegroundColor Cyan
Write-Host "  Root Folder: $rootFolder"
Write-Host "  AD Group: $adGroup"
Write-Host "  Script Path: $scriptPath"
Write-Host "  Error Log: $errorLogFile"

# === FUNCTION TO GRANT PERMISSIONS ===
function Grant-ReadAccessToAdGroup {
    param (
        [string]$path,
        [string]$adGroup
    )

    $permission = "ReadAndExecute"
    $inheritanceFlag = [System.Security.AccessControl.InheritanceFlags]::ContainerInherit, [System.Security.AccessControl.InheritanceFlags]::ObjectInherit
    $propagationFlag = [System.Security.AccessControl.PropagationFlags]::None
    $rule = New-Object System.Security.AccessControl.FileSystemAccessRule($adGroup, $permission, $inheritanceFlag, $propagationFlag, "Allow")

    try {
        $acl = Get-Acl $path
        $existingRule = $acl.Access | Where-Object {
            $_.IdentityReference -eq $adGroup -and
            $_.FileSystemRights -eq [System.Security.AccessControl.FileSystemRights]::ReadAndExecute -and
            $_.IsInherited -eq $true
        }

        if ($existingRule) {
            Write-Output "Permissions for $adGroup are already inherited on $path. Skipping..."
        } else {
            Write-Output "Adding inherited Read & Execute permissions for $adGroup on $path"
            $acl.AddAccessRule($rule)
            Set-Acl -Path $path -AclObject $acl
        }

        # Recurse into child directories
        Get-ChildItem -Path $path -Recurse -Directory | ForEach-Object {
            try {
                $childAcl = Get-Acl $_.FullName
                $childExistingRule = $childAcl.Access | Where-Object {
                    $_.IdentityReference -eq $adGroup -and
                    $_.FileSystemRights -eq [System.Security.AccessControl.FileSystemRights]::ReadAndExecute -and
                    $_.IsInherited -eq $true
                }

                if ($childExistingRule) {
                    Write-Output "Permissions already inherited on $($_.FullName). Skipping..."
                } else {
                    Write-Output "Adding permissions for $adGroup on $($_.FullName)"
                    $childAcl.AddAccessRule($rule)
                    Set-Acl -Path $_.FullName -AclObject $childAcl
                }
            } catch {
                $errorMsg = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Error processing $($_.FullName): $($_.Exception.Message)"
                $errorMsg | Out-File -Append -FilePath $errorLogFile
            }
        }
    } catch {
        $errorMsg = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] General error on $path: $($_.Exception.Message)"
        $errorMsg | Out-File -Append -FilePath $errorLogFile
    }
}

# === EXECUTE ===
Write-Output "Starting permission grant operation..."
Grant-ReadAccessToAdGroup -path $rootFolder -adGroup $adGroup
Write-Output "Process completed. Errors, if any, logged to $errorLogFile"

# End transcript
Stop-Transcript
