# Install the Exchange Online PowerShell module if you haven't already
# Uncomment the line below if you need to install the module
# Install-Module -Name ExchangeOnlineManagement

# Import the Exchange Online PowerShell module
Import-Module ExchangeOnlineManagement

# Connect to Exchange Online (you will be prompted for credentials)
Connect-ExchangeOnline

# Get all user mailbox properties, including the proxy addresses
$mailboxes = Get-Mailbox -ResultSize Unlimited

# Create an array to store mailbox information
$proxyAddressesInfo = @()

# Loop through each mailbox and extract the proxy addresses
foreach ($mailbox in $mailboxes) {
    $userProxyAddresses = @()

    # Extract the primary SMTP address and add it to the list of proxy addresses
    $userProxyAddresses += $mailbox.PrimarySmtpAddress.ToString()

    # Extract all additional proxy addresses (aliases)
    foreach ($proxy in $mailbox.EmailAddresses) {
        if ($proxy.PrefixString -eq "smtp") {
            $userProxyAddresses += $proxy.SmtpAddress
        }
    }

    # Add mailbox information to the array
    $mailboxInfo = @{
        "User" = $mailbox.UserPrincipalName
        "ProxyAddresses" = $userProxyAddresses -join ";"
    }
    $proxyAddressesInfo += New-Object PSObject -Property $mailboxInfo
}

# Output the mailbox information to a CSV file
$csvFilePath = "C:\Scripts\ProxyAddresses.csv"
$proxyAddressesInfo | Export-Csv -Path $csvFilePath -NoTypeInformation

# Disconnect from Exchange Online
Disconnect-ExchangeOnline

Write-Host "Proxy addresses have been exported to $csvFilePath."
