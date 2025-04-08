<#
.SYNOPSIS
Checks for updates to all non-Microsoft PowerShell modules and installs them.

.DESCRIPTION
This script retrieves all locally installed PowerShell modules (excluding Microsoft.*),
checks for available updates, installs newer versions, and optionally removes older versions.
It uses PowerShellGet and requires administrative rights for system-scoped modules.

.PARAMETER RemoveOldVersions
If specified, the script will uninstall old versions of modules after updates are complete.

.PARAMETER Exclude
A list of module names to exclude from updating (e.g., modules with custom modifications).

.EXAMPLE
.\Update-InstalledModules.ps1 -RemoveOldVersions

.EXAMPLE
.\Update-InstalledModules.ps1 -Exclude "Posh-Git", "OhMyPosh"

.NOTES
- Requires PowerShell 5.0+ (PowerShellGet module)
- Requires an internet connection
- Must be run as administrator to update system-installed modules
#>

[CmdletBinding()]
param (
    [switch]$RemoveOldVersions,
    
    [string[]]$Exclude = @()
)

# Ensure PowerShellGet is available
if (-not (Get-Module -ListAvailable -Name PowerShellGet)) {
    Write-Warning "PowerShellGet module is required but not installed. Please install it from https://www.powershellgallery.com/packages/PowerShellGet."
    return
}

# Get distinct non-Microsoft module names
$installedModules = Get-Module -ListAvailable |
    Where-Object { $_.Name -notlike "Microsoft.*" -and $_.Name -notin $Exclude } |
    Select-Object -ExpandProperty Name -Unique

# Track updated modules
$updatedModules = @()

# Update each module
foreach ($moduleName in $installedModules) {
    try {
        Write-Host "Checking for updates to '$moduleName'..." -ForegroundColor Cyan
        Update-Module -Name $moduleName -Force -ErrorAction Stop
        Write-Host "Updated '$moduleName'" -ForegroundColor Green
        $updatedModules += $moduleName
    }
    catch {
        Write-Warning "Could not update module '$moduleName': $($_.Exception.Message)"
    }
}

# Remove old versions if specified
if ($RemoveOldVersions) {
    $groupedModules = Get-Module -ListAvailable |
        Where-Object { $_.Name -notlike "Microsoft.*" -and $_.Name -notin $Exclude } |
        Group-Object -Property Name |
        Where-Object { $_.Count -gt 1 }

    foreach ($group in $groupedModules) {
        $moduleName = $group.Name
        $versions = $group.Group | Select-Object -ExpandProperty Version
        $latestVersion = $versions | Sort-Object -Descending | Select-Object -First 1
        $oldVersions = $versions | Where-Object { $_ -ne $latestVersion }

        foreach ($version in $oldVersions) {
            try {
                Write-Host "Removing old version $version of '$moduleName'..." -ForegroundColor Yellow
                Uninstall-Module -Name $moduleName -RequiredVersion $version -Force -ErrorAction Stop
            }
            catch {
                Write-Warning "Failed to remove version $version of '$moduleName': $($_.Exception.Message)"
            }
        }
    }
}

Write-Host "`nModule update process complete." -ForegroundColor Cyan
if ($updatedModules.Count -gt 0) {
    Write-Host "Modules updated: $($updatedModules -join ', ')" -ForegroundColor Green
} else {
    Write-Host "No modules were updated." -ForegroundColor Gray
}
