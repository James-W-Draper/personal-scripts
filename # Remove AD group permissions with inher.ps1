# Remove AD group permissions with inheritance check
# Version 3.0, Created 11092024

# Define variables for permissions and folder path
$rootFolder = "J:\F_Drive\Users$" # Replace with the top-level folder path
$adGroup = "cwglobal\SD_Admins_L2"  # Replace with your AD group name
$scriptPath = "C:\scripts" # Replace with folder location for reports

# Define variables for error logging
$date = Get-Date -Format "ddMMyyyy"
$datetime = get-date -format "ddMMyyyy HH:mm"
$serverName = $env:COMPUTERNAME

# Replace backslashes and colons in $rootFolder with underscores directly for reporting purposes
$rootFolder = $rootFolder -replace '[\\:]', '_'
$rootFolder = $rootFolder -replace '__+', '_'

# Construct the error log file name including the root folder, server name, and date
$errorLogFile = "${scriptPath}\Errors_${serverName}_${rootFolder}_${date}.txt"

# Output to verify the constructed file path
Write-Output $errorLogFile

# Start transcription of permissions removal
Start-transcript "${scriptPath}\Transcript_${serverName}_${rootFolder}_${date}.txt"

# Add a line to the transcript with the key variables
Write-Host "Script started at ${datetime} with the following variables:" >> "${scriptPath}\Transcript_${serverName}_${rootFolder}_${adGroup}_${date}.txt"
Write-Host "  RootFolder: $rootFolder" >> "${scriptPath}\Transcript_${serverName}_${rootFolder}_${adGroup}_${date}.txt"
Write-Host "  ADGroup: $adGroup" >> "${scriptPath}\Transcript_${serverName}_${rootFolder}_${adGroup}_${date}.txt"
Write-Host "  ScriptPath: $scriptPath" >> "${scriptPath}\Transcript_${serverName}_${rootFolder}_${adGroup}_${date}.txt"
Write-Host "  Error Log: $errorLogFile" >> "${scriptPath}\Transcript_${serverName}_${rootFolder}_${date}.txt"


# Function to remove read access for the AD group from all folders under $rootFolder
function Remove-ReadAccessFromAdGroup {
  param (
    [string]$path,
    [string]$adGroup
  )

  # Check if the AD group has read access with inheritance
  $acl = Get-Acl $path
  $existingRule = $acl.Access | Where-Object {
    $_.IdentityReference -eq $adGroup -and $_.FileSystemRights -eq [System.Security.AccessControl.FileSystemRights]::ReadAndExecute -and $_.IsInherited -eq $true
  }

  if ($existingRule) {
    Write-Output "Removing permissions for $adGroup from $path (with inheritance)"
    $acl.RemoveAccessRule($existingRule)
    Set-Acl -Path $path -AclObject $acl
  } else {
    Write-Output "Permissions for $adGroup were not found on $path. Skipping..."
  }

  # Recurse through child directories
  Get-ChildItem -Path $path -Recurse -Directory | ForEach-Object {
    try {
      if ($null -ne $_) {
        $childAcl = Get-Acl $_.FullName
        $childExistingRule = $childAcl.Access | Where-Object {
          $_.IdentityReference -eq $adGroup -and $_.FileSystemRights -eq [System.Security.AccessControl.FileSystemRights]::ReadAndExecute -and $_.IsInherited -eq $true
        }

        if ($childExistingRule) {
          Write-Output "Removing permissions for $adGroup from $($_.FullName) (with inheritance)"
          $childAcl.RemoveAccessRule($childExistingRule)
          Set-Acl -Path $_.FullName -AclObject $childAcl
        } else {
          Write-Output "Permissions for $adGroup were not found on $($_.FullName). Skipping..."
        }
      }
    } catch {
      # Log errors for individual directories
      $errorMsg = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Error processing $_.FullName: $($_.Exception.Message)"
      $errorMsg | Out-File -Append -FilePath $errorLogFile
    }
  }
}

# Remove read access from the AD group
Write-Output "Removing read access for the AD group..."
Remove-ReadAccessFromAdGroup -path $rootFolder -adGroup $adGroup