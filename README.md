# PowerShell Scripts for Microsoft 365 and Active Directory Management

This repository contains a collection of PowerShell scripts designed to assist administrators in managing, auditing, and automating tasks within Microsoft 365 and Active Directory environments. Whether you're handling user accounts, mailboxes, permissions, or generating reports, these scripts aim to streamline your workflows.

---

## Usage

### Prerequisites
- **PowerShell Version:** Scripts are generally tested with PowerShell 5.1 and later. However, newer versions (like PowerShell 7.x) are recommended for cross-platform compatibility and modern features.
- **Execution Policy:** You might need to set the PowerShell execution policy to run scripts. For example, `Set-ExecutionPolicy RemoteSigned`. For more details, refer to the official Microsoft documentation: [about_Execution_Policies](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_execution_policies).
- **Modules:** Scripts may require specific PowerShell modules to be installed (e.g., for Azure AD, Exchange Online, SharePoint Online). Check script comments or error messages for required modules and install them using `Install-Module <ModuleName>`.
- **Permissions:** Ensure you have appropriate administrative permissions in Active Directory, Microsoft 365 services (Exchange, SharePoint, Teams, Azure AD), and on the local machine to run these scripts effectively.

### Connecting to Services
- **Microsoft Graph / Azure AD:** Many modern M365 scripts might use the Microsoft Graph SDK or Azure AD PowerShell modules (e.g., AzureAD, MSOnline). The `Connect-MgGraphWithCertificate.ps1` script in this repository provides an example for certificate-based authentication with Microsoft Graph.
- **Exchange Online:** Typically, scripts connect to Exchange Online using `Connect-ExchangeOnline`.
- **SharePoint Online:** Connection to SharePoint Online is often established using `Connect-PnPOnline` (for PnP PowerShell) or `Connect-SPOService` (for SharePoint Online Management Shell).
- **Active Directory:** Scripts targeting on-premises Active Directory generally need to be run on a domain-joined machine or a machine with AD Remote Server Administration Tools (RSAT) installed. They will use your current domain credentials or may prompt for them.

### Running Scripts
- **General Execution:** To run a script, navigate to its directory in PowerShell and execute it using its path. For example: `.\ScriptName.ps1 -ParameterName "ParameterValue"`
- **Parameters and Help:** Always check individual scripts for specific parameters and usage instructions. This information is often found in comments at the beginning of the script or can be accessed using the `Get-Help` cmdlet (e.g., `Get-Help .\ScriptName.ps1 -Full`).

## Scripts Overview

### User & Group Management (Active Directory & Azure AD/Entra ID)
- **Add-AdUserToGroupByDisplayName.ps1:** Adds an Active Directory user to a group based on the group's display name.
- **Get-ExpiredADUsers.ps1:** Retrieves a list of expired user accounts from Active Directory.
- **Remove-ADUserFromGroupByDisplayName.ps1:** Removes an Active Directory user from a group based on the group's display name.
- **Revert-ADGroupMoves.ps1:** Reverts changes made to Active Directory group memberships.
- **Simulate-GroupRestorationFromExcel.ps1:** Simulates the restoration of group memberships from an Excel file.
- **Sync-ManagersToAdGroup.ps1:** Synchronizes managers from user profiles to an Active Directory group.
- **Update-ADGroupWithManagers.ps1:** Updates an Active Directory group with managers from user profiles.
- **get-addirectreports.ps1:** Retrieves direct reports for a user from ActiveDirectory.

### Mailbox Management (Exchange Online)
- **Convert-MailboxesToShared.ps1:** Converts user mailboxes to shared mailboxes in Exchange Online.
- **Export-MailTrafficStats.ps1:** Exports statistics related to mail traffic from Exchange Online.
- **Export-MailboxFolderStatsWithArchive.ps1:** Exports folder statistics for mailboxes, including archives.
- **Export-ProxyAddresses.ps1:** Exports proxy addresses (email aliases) for mailboxes.
- **Get-MailboxAuditTrail.ps1:** Retrieves audit trail logs for a specified mailbox.
- **Get-MailboxAutoReplyStatus.ps1:** Checks the auto-reply (out-of-office) status for mailboxes.
- **Get-MailboxSizeReport.ps1:** Generates a report of mailbox sizes in Exchange Online.
- **Get-NonOwnerMailboxAuditReport.ps1:** Generates a report of non-owner access to mailboxes.
- **remove-mailboxpermission.ps1:** Removes permissions from a mailbox.
- **Set-SharedMailboxesForOU.ps1:** Configures shared mailboxes for a specific organizational unit (OU).

### Permissions Management (AD & M365)
- **Export-CalendarPermissions.ps1.ps1:** Exports calendar permissions for users.
- **Export-NTFSPermissionsReport.ps1:** Generates a report of NTFS permissions for files and folders.
- **Get-MailboxPermissionSummary.ps1:** Provides a summary of mailbox permissions.
- **Get-MailboxPermissionsReportGrouped.ps1:** Generates a grouped report of mailbox permissions.
- **Grant-RoomAccessPermissions.ps1:** Grants access permissions to room mailboxes.
- **Remove-StaleMailboxPermissions.ps1:** Removes stale or outdated mailbox permissions.
- **Set-AdGroupReadAccessFromCsv.ps1:** Sets read access permissions for Active Directory groups based on a CSV file.
- **Set-AdGroupReadAccessRecursive.ps1:** Sets recursive read access permissions for Active Directory groups.
- **Set-NTFSPermissionsAndOwnership.ps1:** Sets NTFS permissions and ownership for files and folders.
- **remove-readexecutepermissions.ps1:** Removes read and execute permissions from files or folders.

### Reporting & Auditing (M365 & AD)
- **Audit-ExternalUserAccess.ps1:** Audits external user access to resources.
- **Export-GuestUsersTeamsMembershipReport.ps1:** Exports a report of guest users' membership in Microsoft Teams.
- **Export-M365SharedChannelExternalMembers.ps1:** Exports a report of external members in Microsoft 365 shared channels.
- **Export-M365SharedChannelsGuestMembers.ps1:** Exports a report of guest members in Microsoft 365 shared channels.
- **Export-NonOwnerMailboxAccessReport.ps1:** Generates a report of non-owner access to mailboxes.
- **Export-TeamsReports.ps1:** Exports various reports related to Microsoft Teams.
- **Export-UnifiedGroupOwners.ps1:** Exports a list of owners for Microsoft 365 Unified Groups.
- **Export-guestUsersReport.ps1:** Generates a report of guest users in the environment.
- **Export-mfastatusReport.ps1:** Generates a report of multi-factor authentication (MFA) status for users.
- **Get-ContentSearchFolderTargets.ps1:** Retrieves folder targets for content searches in Microsoft 365.
- **Get-CopilotAuditLogs.ps1:** Retrieves audit logs related to Copilot usage.
- **Get-CopilotAuditlog.ps1:** Retrieves audit logs related to Copilot usage.
- **Get-LatestCalendarEventPerUser.ps1:** Retrieves the latest calendar event for each user.
- **get-copilotcount.ps1:** Counts Copilot related entities or activities.

### File & Site Management (SharePoint Online/OneDrive)
- **Delete-SiteFilesAndLog.ps1:** Deletes files from a SharePoint site and logs the actions.
- **Remove-FilesFromCsv.ps1:** Removes files listed in a CSV file from a SharePoint site or OneDrive.
- **Remove-FilesWithAudit.ps1:** Removes files with auditing from a SharePoint site or OneDrive.
- **Remove-FilesWithAuditAndReport.ps1:** Removes files with auditing and generates a report.
- **Set-SPOAppSitePermissions.ps1:** Sets app permissions for a SharePoint Online site.

### License Management (Microsoft 365)
- **Manage-M365Licenses.ps1:** Manages Microsoft 365 licenses for users.

### Azure AD / Entra ID Management
- **Connect-MgGraphWithCertificate.ps1:** Connects to Microsoft Graph API using certificate-based authentication.
- **Get-AADPrivilegedAccounts.ps1:** Retrieves privileged accounts from Azure Active Directory.
- **Get-ExpiringAADAppCredentials.ps1:** Retrieves Azure Active Directory application credentials that are expiring soon.
- **New-EntraAppRegistration.ps1:** Creates a new application registration in Entra ID (Azure AD).

### Utility & Miscellaneous Scripts
- **Mail test.ps1:** Likely a script for testing email functionality.
- **Update-InstalledModules.ps1.ps1:** Updates installed PowerShell modules.
- **termsrv_rdp_patch.ps1:** Applies a patch related to Terminal Services or Remote Desktop Protocol.

## Contributing

Contributions to this collection of scripts are welcome! If you have a script that you believe would be a good addition, or if you have improvements or bug fixes for existing scripts, please follow these general guidelines:

1.  **Fork the repository.**
2.  **Create a new branch** for your feature or bug fix (e.g., `feature/your-feature-name` or `fix/script-bug-fix`).
3.  **Write clear and concise commit messages.**
4.  **Ensure your script is well-commented**, explaining its purpose, parameters, and any complex logic.
5.  **Test your script thoroughly** in a non-production environment.
6.  **Consider cross-platform compatibility** if applicable (PowerShell 7+).
7.  **Submit a pull request** to the `main` branch of this repository. In your pull request, please describe the changes you've made and the problem they solve.

While there are no strict coding style rules yet, try to maintain consistency with the existing scripts in terms of naming conventions and formatting.

### Formatting guidance
[basic-writing-and-formatting-syntax](https://github.com/github/docs/blob/main/content/get-started/writing-on-github/getting-started-with-writing-and-formatting-on-github/basic-writing-and-formatting-syntax.md)

## License

This project is licensed under the MIT License.

**MIT License**

Copyright (c) [Year] Project Contributors

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
