function Get-ADDirectReports {
    <#
    .SYNOPSIS
        Retrieves direct reports from Active Directory for one or more specified users.

    .DESCRIPTION
        This function queries Active Directory for users who report directly or indirectly
        (when -Recurse is used) to the specified Identity. It returns detailed information
        about each reporting user, including their name, mail, and manager.

    .PARAMETER Identity
        One or more user accounts (e.g. sAMAccountName, DistinguishedName, etc.) to inspect.

    .PARAMETER Recurse
        If specified, recursively retrieves all indirect reports as well.

    .EXAMPLE
        Get-ADDirectReports -Identity "j.smith"

        Retrieves users who directly report to j.smith.

    .EXAMPLE
        Get-ADDirectReports -Identity "j.smith" -Recurse

        Retrieves both direct and indirect reports to j.smith.

    .NOTES
        Author: Francois-Xavier Cat (Original)
        Adapted and modernised by: James Draper
        Source: https://lazywinadmin.com/2014/10/powershell-who-reports-to-whom-active.html

    .LINK
        https://github.com/lazywinadmin/PowerShell
    #>

    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string[]]$Identity,

        [Parameter()]
        [switch]$Recurse
    )

    begin {
        try {
            if (-not (Get-Module -Name ActiveDirectory)) {
                Import-Module -Name ActiveDirectory -ErrorAction Stop -Verbose:$false
            }
        } catch {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }

    process {
        foreach ($User in $Identity) {
            try {
                if ($Recurse) {
                    Write-Verbose "Processing $User recursively"
                    Get-ADUser -Identity $User -Properties DirectReports | ForEach-Object {
                        $_.DirectReports | ForEach-Object {
                            Get-ADUser -Identity $_ -Properties * |
                                Select-Object -Property *, @{Name='ManagerAccount'; Expression={ (Get-ADUser -Identity $_.Manager).SamAccountName }}

                            # Recursive call
                            Get-ADDirectReports -Identity $_ -Recurse
                        }
                    }
                } else {
                    Write-Verbose "Processing $User"
                    Get-ADUser -Identity $User -Properties DirectReports |
                        Select-Object -ExpandProperty DirectReports |
                        ForEach-Object {
                            Get-ADUser -Identity $_ -Properties * |
                                Select-Object -Property *, @{Name='ManagerAccount'; Expression={ (Get-ADUser -Identity $_.Manager).SamAccountName }}
                        }
                }
            } catch {
                $PSCmdlet.ThrowTerminatingError($_)
            }
        }
    }

    end {
        Remove-Module -Name ActiveDirectory -ErrorAction SilentlyContinue -Verbose:$false | Out-Null
    }
}
