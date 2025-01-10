# Get a list of all installed modules, check for updates, and install them.  This is a good way to keep your modules up to date. This script will not update modules that are installed in the user's profile.
# Requires PowerShell 3.0 or later.
# Requires the PowerShellGet module, which is included with PowerShell 5.0 and later, or available for download at https://www.powershellgallery.com/packages/PowerShellGet.
# Requires an internet connection.
# Requires local administrator rights.

# Get a list of all installed modules.
$Modules = Get-Module -ListAvailable | Where-Object {$_.Name -notlike 'Microsoft.*'} | Select-Object -ExpandProperty Name

# Check for updates to each module, and install them.
foreach ($Module in $Modules) {
    Write-Host "Checking for updates to $Module..."
    Update-Module -Name $Module -Force
}

# Check for multiple versions of the same module and remove the older versions
Get-Module -ListAvailable | Where-Object {$_.Name -notlike 'Microsoft.*'} | Group-Object -Property Name | Where-Object {$_.Count -gt 1} | ForEach-Object {
    $Module = $_.Name
    $Versions = $_.Group | Select-Object -ExpandProperty Version
    $LatestVersion = $Versions | Sort-Object -Descending | Select-Object -First 1
    $OldVersions = $Versions | Where-Object {$_ -ne $LatestVersion}
    Write-Host "Removing $($OldVersions.Count) old versions of $Module..."
    foreach ($OldVersion in $OldVersions) {
        Uninstall-Module -Name $Module -RequiredVersion $OldVersion -Force
    }
}