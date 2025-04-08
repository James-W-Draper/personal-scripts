<#
.SYNOPSIS
Recursively sets NTFS permissions and ownership on a specified folder and its contents.

.DESCRIPTION
This script sets Read and Execute permissions for a specified Active Directory group on a given folder,
its subfolders, and files. It also sets ownership to a specified local or domain group (e.g., Administrators).
All changes and outcomes are logged to a CSV report.

.PARAMETER TargetFolder
The full path of the folder where permissions and ownership will be set.

.PARAMETER ADGroup
The Active Directory group to which Read and Execute permissions will be granted.

.PARAMETER NewOwner
The account (local or domain group) to assign as the owner of each file/folder.

.PARAMETER ReportPath
The full path to the CSV file where the permission/ownership change log will be saved.

.EXAMPLE
.\Set-NTFSPermissionsAndOwnership.ps1 -TargetFolder "F:\Data" -ADGroup "domain\ad_group" -NewOwner "BUILTIN\Administrators" -ReportPath "C:\scripts\PermissionsLog.csv"
#>

param (
    [Parameter(Mandatory = $true)]
    [string]$TargetFolder,

    [Parameter(Mandatory = $true)]
    [string]$ADGroup,

    [Parameter(Mandatory = $true)]
    [string]$NewOwner,

    [Parameter(Mandatory = $true)]
    [string]$ReportPath
)

# Create the CSV header
Remove-Item -Path $ReportPath -ErrorAction Ignore
Add-Content -Path $ReportPath -Value '"FilePath","Status","Details"' + [Environment]::NewLine

function Write-Log {
    param (
        [string]$Path,
        [string]$Status,
        [string]$Details
    )
    $entry = '"{0}","{1}","{2}"' -f $Path, $Status, $Details
    Add-Content -Path $ReportPath -Value $entry
}

# Define permissions
$folderRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
    $ADGroup,
    "ReadAndExecute",
    [System.Security.AccessControl.InheritanceFlags]::ContainerInherit,
    [System.Security.AccessControl.PropagationFlags]::None,
    [System.Security.AccessControl.AccessControlType]::Allow
)

$fileRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
    $ADGroup,
    "ReadAndExecute",
    [System.Security.AccessControl.InheritanceFlags]::ObjectInherit,
    [System.Security.AccessControl.PropagationFlags]::None,
    [System.Security.AccessControl.AccessControlType]::Allow
)

function Set-Ownership {
    param ([string]$Path)
    try {
        $acl = Get-Acl -Path $Path
        $acl.SetOwner([System.Security.Principal.NTAccount]$NewOwner)
        Set-Acl -Path $Path -AclObject $acl
        Write-Log -Path $Path -Status "Ownership Set" -Details "Ownership successfully updated."
    } catch {
        Write-Log -Path $Path -Status "Ownership Failed" -Details $_.Exception.Message
    }
}

function Set-Permissions {
    param ([string]$Path)
    try {
        $acl = Get-Acl -Path $Path
        $acl.AddAccessRule($folderRule)
        $acl.AddAccessRule($fileRule)
        Set-Acl -Path $Path -AclObject $acl
        Write-Log -Path $Path -Status "Permissions Set" -Details "Permissions successfully updated."
    } catch {
        Write-Log -Path $Path -Status "Permissions Failed" -Details $_.Exception.Message
    }
}

# Process root folder
Write-Log -Path $TargetFolder -Status "Processing Root Folder" -Details ""
Set-Ownership -Path $TargetFolder
Set-Permissions -Path $TargetFolder

# Process subfolders/files recursively
Get-ChildItem -Path $TargetFolder -Recurse -Force | ForEach-Object {
    $itemPath = $_.FullName
    Write-Log -Path $itemPath -Status "Processing Item" -Details ""
    Set-Ownership -Path $itemPath
    Set-Permissions -Path $itemPath
}

Write-Output "Processing complete. Report available at: $ReportPath"
