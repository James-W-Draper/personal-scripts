Connect-ExchangeOnline
Connect-MgGraph
Select-MgProfile beta
Connect-MicrosoftTeams

$ReportUsers = Get-MgUser -Filter "userType eq 'Guest'" -ConsistencyLevel eventual -All -Property @(
    'UserPrincipalName'
    'SignInActivity'
    'CreatedDateTime'
    'DisplayName'
    'Mail'
    'OnPremisesImmutableId'
    'OnPremisesDistinguishedName'
    'OnPremisesLastSyncDateTime'
    'SignInSessionsValidFromDateTime'
    'RefreshTokensValidFromDateTime'
    'id'
) | Select-Object @(
    'UserPrincipalName'
    'CreatedDateTime'
    'DisplayName'
    'Mail'
    'SignInSessionsValidFromDateTime'
    'RefreshTokensValidFromDateTime'
    'id'
    @{n='LastSignInDateTime'; e={[datetime]$_.SignInActivity.LastSignInDateTime}}
) 
| ForEach-Object {
    $user = $_
    $teamsMembership = Get-TeamUser -User $user.UserPrincipalName
    $teamsMembership | ForEach-Object {
        [PSCustomObject]@{
            UserPrincipalName = $user.UserPrincipalName
            CreatedDateTime = $user.CreatedDateTime
            DisplayName = $user.DisplayName
            Mail = $user.Mail
            SignInSessionsValidFromDateTime = $user.SignInSessionsValidFromDateTime
            RefreshTokensValidFromDateTime = $user.RefreshTokensValidFromDateTime
            LastSignInDateTime = $user.LastSignInDateTime
            lastNonInteractiveSignInDateTime = $user.lastNonInteractiveSignInDateTime
            TeamDisplayName = $_.DisplayName
        }
    }
}
$Common_ExportExcelParams = @{
    BoldTopRow   = $true
    AutoSize     = $true
    AutoFilter   = $true
    FreezeTopRow = $true
}

$FileDate = Get-Date -Format yyyyMMddTHHmmss

$ReportUsers | Sort-Object UserPrincipalName | Export-Excel @Common_ExportExcelParams -Path ("c:\scripts\" + $FileDate + "_report.xlsx") -WorksheetName report
