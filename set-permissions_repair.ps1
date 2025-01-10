# DONE UKS-PRD-FIL-01 UKRDCDAT01 - G:\Drive_X\InigoData
# DONE UKS-PRD-FIL-01 UKRDCDAT01 - G:\Drive_G\ArchiveData 
# DONE UKS-PRD-FIL-01 UKRDCDAT01 - G:\Drive_H\Data (Access issues on majority of the folders within Data) 
# DONE UKS-PRD-FIL-01 LSGPRDFS01A - E:\Drive_L
# DONE UKS-PRD-FIL-01 LSGPRDFS01A - E:\Drive_G\Data\Enstar O Drive Data – To Be Deleted On Sept_1_2014
# DONE UKS-PRD-FIL-01 LSGPRDFS01A - E:\Drive_G\Data\Shared\File System
# DONE UKS-PRD-FIL-01 LSGPRDFS01A - E:\Drive_G\Data\Shared\Ioanna
# DONE UKS-PRD-FIL-01 LSGPRDFS01A - E:\Drive_G\Data\Shared\TBIL - MGA

# DONE USE2-PRD-FIL-01 Usrdcdat02 - F:\H_drive

# DONE USE2-PRD-FIL-01 PRDFIL01 - K:\D_Drive\DFSData\DeptsData\Accounting
# DONE UKS-PRD-FIL-01 UKRDCDAT01 G:\Drive_H\Data

# DONE BUT BUGS USE2-PRD-FIL-01 Usrdcdat02 F:\F_Drive\USWARDC10
# USE2-PRD-FIL-01 Usrdcdat02 - F:\F_Drive\USSTPDC08

# UKS-PRD-FIL-01 LSGPRDFS01A - E:\Drive_G\


# Define the folder path, AD group, and local group for ownership
$FolderPath = "F:\F_Drive\USWARDC10"
$ADGroup = "Cwglobal\sd_admins_l2"
$LocalGroup = "BUILTIN\Administrators"  # Use BUILTIN\Administrators for the local Administrators group
$CsvLogFilePath = "C:\scripts\Permissions_Report_archivedata.csv"

# Clear any previous log file and add headers
Remove-Item -Path $CsvLogFilePath -ErrorAction Ignore
$headerLine = '"FilePath","Status","Details"' + [Environment]::NewLine
Add-Content -Path $CsvLogFilePath -Value $headerLine


# Function to write results to CSV with quoted fields
function Write-Result {
    param (
        [string]$path,
        [string]$status,
        [string]$details
    )
    # Construct the CSV line as a single string with quoted fields
    $logLine = '"' + $path + '","' + $status + '","' + $details + '"'
    Add-Content -Path $CsvLogFilePath -Value $logLine
}

# Define permissions (Read & Execute)
$permission = "ReadAndExecute"

# Create separate access rules for folders and files
$folderAccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
    $ADGroup, 
    $permission, 
    [System.Security.AccessControl.InheritanceFlags]::ContainerInherit, 
    [System.Security.AccessControl.PropagationFlags]::None, 
    [System.Security.AccessControl.AccessControlType]::Allow
)

$fileAccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
    $ADGroup, 
    $permission, 
    [System.Security.AccessControl.InheritanceFlags]::ObjectInherit, 
    [System.Security.AccessControl.PropagationFlags]::None, 
    [System.Security.AccessControl.AccessControlType]::Allow
)

# Function to set ownership with logging
function Set-Ownership {
    param (
        [string]$path,
        [string]$owner
    )
    try {
        $acl = Get-Acl -Path $path
        $acl.SetOwner([System.Security.Principal.NTAccount]$owner)
        Set-Acl -Path $path -AclObject $acl
        Write-Result -path $path -status "Ownership Set" -details "Ownership set successfully"
    } catch {
        Write-Result -path $path -status "Ownership Failed" -details $_.Exception.Message
    }
}

# Function to set permissions with logging
function Set-Permissions {
    param (
        [string]$path
    )
    try {
        $itemAcl = Get-Acl -Path $path
        $itemAcl.AddAccessRule($folderAccessRule)
        $itemAcl.AddAccessRule($fileAccessRule)
        Set-Acl -Path $path -AclObject $itemAcl
        Write-Result -path $path -status "Permissions Set" -details "Permissions set successfully"
    } catch {
        Write-Result -path $path -status "Permissions Failed" -details $_.Exception.Message
    }
}

# First, set ownership and log the result for the top-level folder
Write-Result -path $FolderPath -status "Processing Top-Level Folder" -details ""
Set-Ownership -path $FolderPath -owner $LocalGroup
Set-Permissions -path $FolderPath

# Recursively apply ownership and permissions to all items within the folder
Get-ChildItem -Path $FolderPath -Recurse | ForEach-Object {
    $itemPath = $_.FullName
    Write-Result -path $itemPath -status "Processing Item" -details ""
    Set-Ownership -path $itemPath -owner $LocalGroup
    Set-Permissions -path $itemPath
}

Write-Output "Ownership and permissions process completed. Report generated at $CsvLogFilePath."
