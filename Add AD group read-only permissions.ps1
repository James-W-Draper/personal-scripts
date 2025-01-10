# Add AD group read-only permissions with inheritance check
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

# Start transcription of permissions
Start-transcript "${scriptPath}\Transcript_${serverName}_${rootFolder}_${date}.txt"

# Add a line to the transcript with the key variables
Write-Host "Script started at ${datetime} with the following variables:" >> "${scriptPath}\Transcript_${serverName}_${rootFolder}_${adGroup}_${date}.txt"
Write-Host "  RootFolder: $rootFolder" >> "${scriptPath}\Transcript_${serverName}_${rootFolder}_${adGroup}_${date}.txt"
Write-Host "  ADGroup: $adGroup" >> "${scriptPath}\Transcript_${serverName}_${rootFolder}_${adGroup}_${date}.txt"
Write-Host "  ScriptPath: $scriptPath" >> "${scriptPath}\Transcript_${serverName}_${rootFolder}_${adGroup}_${date}.txt"
Write-Host "  Error Log: $errorLogFile" >> "${scriptPath}\Transcript_${serverName}_${rootFolder}_${date}.txt"


# Function to grant read access to all folders under $rootFolder
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
    # Get ACL for current path
    $acl = Get-Acl $path

    # Check if the AD group already has read access with inheritance
    $existingRule = $acl.Access | Where-Object {
      $_.IdentityReference -eq $adGroup -and $_.FileSystemRights -eq [System.Security.AccessControl.FileSystemRights]::ReadAndExecute -and $_.IsInherited -eq $true
    }

    if ($existingRule) {
      Write-Output "Permissions for $adGroup are already inherited for $path. Skipping..."
    } else {
      Write-Output "Adding permissions for $adGroup to $path (with inheritance)"
      $acl.AddAccessRule($rule)
      Set-Acl -Path $path -AclObject $acl
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
            Write-Output "Permissions for $adGroup are already inherited for $($_.FullName). Skipping..."
          } else {
            Write-Output "Adding permissions for $adGroup to $($_.FullName) (with inheritance)"
            $childAcl.AddAccessRule($rule)
            Set-Acl -Path $_.FullName -AclObject $childAcl
          }
        }
      } catch {
        # Log errors for individual directories
        $errorMsg = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Error processing $_.FullName: $($_.Exception.Message)"
        $errorMsg | Out-File -Append -FilePath $errorLogFile
      }
    }
  } catch {
    # Log general errors
    $errorMsg = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] General error: $($_.Exception.Message)"
    $errorMsg | Out-File -Append -FilePath $errorLogFile
  }
}

# Grant read access to the AD group
Write-Output "Granting read access to the AD group..."
Grant-ReadAccessToAdGroup -path $rootFolder -adGroup $adGroup

Write-Output "Process completed. Any errors encountered have been saved to $errorLogFile"
stop-transcript