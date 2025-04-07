<#
.SYNOPSIS
Removes inherited "Read and Execute" file system permissions for a specified Active Directory group from a given root folder and all its subdirectories.

.DESCRIPTION
This script is designed for situations where a specific AD group (e.g., support staff or a decommissioned team) has inherited file permissions across a large directory structure and needs to have those permissions removed in a safe and logged manner.

The script:
- Recursively processes the specified root folder
- Identifies inherited "Read and Execute" permissions for the AD group
- Removes only inherited permissions (explicit permissions are left intact)
- Logs any errors and actions taken to a designated scripts folder
- Records a transcript of the full session for audit or troubleshooting

.PARAMETER rootFolder
The top-level directory to start processing from. All subfolders will also be evaluated.

.PARAMETER adGroup
The AD group (in DOMAIN\GroupName format) whose inherited permissions should be removed.

.PARAMETER scriptPath
Directory where logs and transcript files will be saved.

.NOTES
- Requires administrative privileges
- User must have permission to modify NTFS ACLs on the target directories
- Designed to be safe and reversible if permissions are managed via group inheritance

.VERSION
3.0 - Cleaned and generalised for cross-tenant or organisation-neutral use

.AUTHOR
James Draper
#>

$rootFolder = "D:\Path\To\RootFolder"  # <-- Replace with the target root folder path
$adGroup = "domain\GroupName"          # <-- Replace with your AD group name
$scriptPath = "C:\Scripts"             # <-- Path for logs and transcripts

# === Script Runtime Metadata ===
$date = Get-Date -Format "yyyyMMdd"
$datetime = Get-Date -Format "yyyy-MM-dd HH:mm"
$serverName = $env:COMPUTERNAME

# Sanitize folder name for safe use in file names
$sanitizedRootFolder = ($rootFolder -replace '[\\:\$]', '_') -replace '__+', '_'

# === Logging ===
$errorLogFile = "${scriptPath}\Errors_${serverName}_${sanitizedRootFolder}_${date}.txt"
$transcriptFile = "${scriptPath}\Transcript_${serverName}_${sanitizedRootFolder}_${date}.txt"

Write-Output "Logging errors to: $errorLogFile"
Write-Output "Logging transcript to: $transcriptFile"

Start-Transcript -Path $transcriptFile

Write-Host "Script started at ${datetime}" -ForegroundColor Cyan
Write-Host "  Root Folder: $rootFolder"
Write-Host "  AD Group: $adGroup"
Write-Host "  Script Path: $scriptPath"
Write-Host "  Error Log: $errorLogFile"

# === Function: Remove Read & Execute Inherited Permissions ===
function Remove-ReadAccessFromAdGroup {
    param (
        [string]$path,
        [string]$adGroup
    )

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

    # Recursively process child directories
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
                    Write-Output "No inherited permissions for $adGroup on $($_.FullName). Skipping..."
                }
            }
        } catch {
            $errorMsg = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Error processing $($_.FullName): $($_.Exception.Message)"
            $errorMsg | Out-File -Append -FilePath $errorLogFile
        }
    }
}

# === Run the Function ===
Write-Output "Initiating permission removal process..."
Remove-ReadAccessFromAdGroup -path $rootFolder -adGroup $adGroup

# === Wrap-Up ===
Stop-Transcript
Write-Host "Script execution completed." -ForegroundColor Green
