$expiredUsers = Get-ADUser -Filter {Enabled -eq $true -and AccountExpirationDate -lt (Get-Date)} -Properties Name, SamAccountName, AccountExpirationDate

$expiredUsers | Where-Object { $_.AccountExpirationDate -ne $null } | 
    Select-Object Name, SamAccountName, @{Name="AccountExpirationDate"; Expression={$_.AccountExpirationDate.ToString("yyyy-MM-dd")}} | 
    Sort-Object AccountExpirationDate
