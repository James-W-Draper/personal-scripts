<#
.SYNOPSIS
Generates a report of all users assigned to Azure AD directory roles (privileged accounts).

.DESCRIPTION
This script connects to Azure AD using the AzureAD module and exports a CSV report listing all users
who are members of any Azure AD directory role, including Global Admin, User Admin, etc.

.PARAMETER OutputPath
Specifies the full path to export the report to (CSV file).

.EXAMPLE
.\Get-AADPrivilegedAccounts.ps1

.EXAMPLE
.\Get-AADPrivilegedAccounts.ps1 -OutputPath "C:\Reports\PrivilegedUsers.csv"
#>

param (
    [string]$OutputPath = "C:\Scripts\AllPrivilegedAccountsReport.csv"
)

# Import the AzureAD module
if (-not (Get-Module -ListAvailable -Name AzureAD)) {
    Write-Host "AzureAD module not found. Please install it with: Install-Module AzureAD"
    exit
}
Import-Module AzureAD

# Connect to Azure AD
try {
    Connect-AzureAD
} catch {
    Write-Error "❌ Failed to connect to Azure AD: $_"
    exit
}

# Retrieve all active roles
$allRoles = Get-AzureADDirectoryRole
$privilegedAccounts = @()

foreach ($role in $allRoles) {
    try {
        $roleMembers = Get-AzureADDirectoryRoleMember -ObjectId $role.ObjectId

        foreach ($member in $roleMembers) {
            $privilegedAccounts += [pscustomobject]@{
                DisplayName       = $member.DisplayName
                UserPrincipalName = $member.UserPrincipalName
                Role              = $role.DisplayName
            }
        }
    } catch {
        Write-Warning "⚠️ Failed to get members for role '$($role.DisplayName)': $_"
    }
}

# Export to CSV
$privilegedAccounts | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
Write-Host "`n✅ All Privileged Accounts Report generated:"
Write-Host "📁 $OutputPath"

# Optional console display grouped by role
Write-Host "`n🧾 Summary (users grouped by role):" -ForegroundColor Cyan
$privilegedAccounts | Group-Object Role | ForEach-Object {
    Write-Host "`n🔐 $($_.Name): $($_.Count) user(s)" -ForegroundColor Yellow
    $_.Group | Format-Table DisplayName, UserPrincipalName -AutoSize
}
