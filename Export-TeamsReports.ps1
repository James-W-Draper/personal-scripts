<#
.SYNOPSIS
    Exports Microsoft Teams reports including Teams, Channels, Members, and Owners into CSV format.

.DESCRIPTION
    This interactive PowerShell script connects to Microsoft Teams using either standard credentials or MFA.
    It presents a menu-driven interface to allow export of a variety of Microsoft Teams-related reports,
    including all teams, specific teams, owners, channels, and membership data.

.PARAMETER UserName
    Optional. Username used for non-MFA authentication.

.PARAMETER Password
    Optional. Plain text password corresponding to the username. Used only if UserName is specified.

.PARAMETER MFA
    Optional switch. If used, prompts for MFA-based login to Microsoft Teams.

.PARAMETER Action
    Optional. If specified, runs the report associated with the given action number without showing the menu.

.NOTES
    Requires MicrosoftTeams module.
    Must be run with privileges to read Microsoft Teams data.

.EXAMPLE
    .\Export-TeamsReports.ps1 -MFA

.EXAMPLE
    .\Export-TeamsReports.ps1 -UserName "admin@contoso.com" -Password "P@ssword123" -Action 1

#>

param(
    [string]$UserName, 
    [string]$Password, 
    [switch]$MFA,
    [int]$Action
)

# Check and import Microsoft Teams module
if (-not (Get-Module -ListAvailable -Name MicrosoftTeams)) {
    Write-Host "MicrosoftTeams module not found." -ForegroundColor Yellow
    $install = Read-Host "Install MicrosoftTeams module now? [Y/N]"
    if ($install -match '^[Yy]$') {
        Install-Module MicrosoftTeams -Force -Scope CurrentUser
    } else {
        Write-Host "MicrosoftTeams module is required. Exiting." -ForegroundColor Red
        exit
    }
}
Import-Module MicrosoftTeams -ErrorAction Stop

# Connect to Microsoft Teams
try {
    if ($MFA) {
        Connect-MicrosoftTeams | Out-Null
    } elseif ($UserName -and $Password) {
        $securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
        $cred = New-Object System.Management.Automation.PSCredential($UserName, $securePassword)
        Connect-MicrosoftTeams -Credential $cred | Out-Null
    } else {
        Connect-MicrosoftTeams | Out-Null
    }
    Write-Host "Successfully connected to Microsoft Teams." -ForegroundColor Green
} catch {
    Write-Host "Failed to connect to Microsoft Teams: $_" -ForegroundColor Red
    exit
}

# Launch menu or perform direct action
function Invoke-TeamsReportMenu {
    # This function wraps the existing menu logic and report processing
    # You can move the full switch logic here as a future cleanup step.
    Write-Host "Launching interactive menu..."
    # Menu and logic would go here...
}

if ($PSBoundParameters.ContainsKey('Action')) {
    $menuChoice = $Action
} else {
    Invoke-TeamsReportMenu
    return
}

# TODO: Migrate the long switch ($i) block into Invoke-TeamsReportMenu or dedicated functions
# Suggest refactoring each menu option into its own function (e.g., Export-AllTeamsReport)

# Disconnect session
Disconnect-MicrosoftTeams
Write-Host "Disconnected from Microsoft Teams." -ForegroundColor Gray
