# Delete_folders.ps1
# Written by James Draper 28/10/2024
# Updated 31/10/2024
# Version 1.2 - rewritten some of the functions to use approved verbs
# Version 1.1 - fixed a minor bug with the icacls command.

# $scriptFolder: This variable points to C:\scripts as the folder where both the log file and the CSV will be stored.
# Write-Log: This function logs messages to the log file, including timestamps to track when events happen.
# Test-Permission: This function checks whether the current user has FullControl permissions on the folder using Get-Acl. If the user has the required permissions, it returns $true; otherwise, it returns $false.
# Set-Ownership: If the current user doesn’t have permissions, the script will attempt to claim ownership of the folder using the takeown and icacls commands. If successful, it logs the success; otherwise, it logs the error and skips further actions on the folder.
# Remove-Folder: The function Remove-Item is used to delete the folder and all its contents (-Recurse -Force flags). If successful, it logs the deletion; otherwise, it logs any encountered errors.
# Invoke-FolderProcessing: The script reads the CSV file using Import-Csv, checks permissions for each folder, claims ownership if necessary, and then attempts to delete the folder. Each action is logged to the log file.
# Make sure you save your CSV file ($csvPath) and the script (delete_folders.ps1) within the folder specified in $scriptFolder
# Run the script with administrative privileges to ensure takeown and icacls can execute.

# Define the folder where the log file and CSV are located
$scriptFolder = "C:\scripts"

# Define the paths for the log file and CSV
$logFile = Join-Path -Path $scriptFolder -ChildPath "deletion_log_CHG0077981_use-prd-prf-01.txt"
$csvPath = Join-Path -Path $scriptFolder -ChildPath "folders_CHG0077981_use-prd-prf-01.csv"

# Function to log messages to the log file
function Write-Log {
    param (
        [string]$message
    )
    # Get the current timestamp and format the message
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $message = "$timestamp - $message"
    
    # Append the message to the log file
    Add-Content -Path $logFile -Value $message
}

# Function to check if the user has the necessary permissions to modify the folder
function Test-Permission {
    param (
        [string]$folderPath
    )
    # Use Get-Acl to retrieve the folder's access control list (ACL)
    $acl = Get-Acl -Path $folderPath -ErrorAction SilentlyContinue
    if ($acl) {
        # Check if the current user has FullControl access rights
        $access = $acl.Access | Where-Object { $_.FileSystemRights -match "FullControl" -and $_.IdentityReference.ToString() -eq [System.Security.Principal.WindowsIdentity]::GetCurrent().Name }
        return $null -ne $access
    }
    return $false
}

# Function to claim ownership of the folder if the user lacks permission
function Set-Ownership {
    param (
        [string]$folderPath
    )
    try {
        # Use the takeown command to claim ownership of the folder
        takeown /f $folderPath /r /d y | Out-Null

        # Use the icacls command to grant FullControl permissions to Administrators
        icacls $folderPath /grant *S-1-5-32-544:"(OI)(CI)F" /T | Out-Null  # S-1-5-32-544 is the SID for the Administrators group

        # Log the successful ownership claim
        Write-Log "Ownership claimed and permissions updated for: $folderPath"
        return $true
    } catch {
        # Log the failure and the error message
        Write-Log "Failed to claim ownership for: $folderPath. Error: $_"
        return $false
    }
}

# Function to delete the folder and its contents
function Remove-Folder {
    param (
        [string]$folderPath
    )
    try {
        # Use Remove-Item to delete the folder recursively and force the deletion
        Remove-Item -Path $folderPath -Recurse -Force -ErrorAction Stop
        
        # Log the successful deletion
        Write-Log "Successfully deleted: $folderPath"
        return $true
    } catch {
        # Log any errors that occurred during the deletion
        Write-Log "Failed to delete $folderPath. Error: $_"
        return $false
    }
}

# Function to process each folder in the CSV file
function Invoke-FolderProcessing {
    # Check if the CSV file exists
    if (-not (Test-Path -Path $csvPath)) {
        Write-Host "CSV file not found at: $csvPath"
        Write-Log "CSV file not found at: $csvPath"
        return
    }

    # Read the CSV file which contains the list of folder paths
    $folders = Import-Csv -Path $csvPath

    # Process each folder in the CSV
    foreach ($folder in $folders) {
        $folderPath = $folder.FolderPath

        # Check if the folder exists
        if (Test-Path -Path $folderPath) {
            Write-Log "Processing folder: $folderPath"

            # Check if the user has the necessary permissions
            if (Test-Permission -folderPath $folderPath) {
                Write-Log "Permission exists for: $folderPath"
            } else {
                # If no permission, attempt to claim ownership
                Write-Log "Permission denied for: $folderPath. Attempting to claim ownership."
                if (-not (Set-Ownership -folderPath $folderPath)) {
                    # If ownership claim fails, skip the deletion
                    Write-Log "Skipping deletion for: $folderPath due to ownership issues."
                    continue
                }
            }

            # Attempt to delete the folder
            Remove-Folder -folderPath $folderPath
        } else {
            # Log if the folder does not exist
            Write-Log "Folder does not exist: $folderPath"
        }
    }

    Write-Host "Process completed. Results have been logged to $logFile."
}

# Call the function to process the folders
Invoke-FolderProcessing
