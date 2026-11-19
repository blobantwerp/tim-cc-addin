# Read-only check of the Exchange Online settings that matter for deploying the add-in.
# Signs in with a device code (you log in yourself in a browser). Changes nothing.
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline -ShowBanner:$false

Write-Host "`n== Add-ins enabled for the organization? ==" -ForegroundColor Cyan
Get-OrganizationConfig | Format-List AppsForOfficeEnabled, OAuthApplications

Write-Host "`n== Existing organization add-ins ==" -ForegroundColor Cyan
Get-App -OrganizationApp | Format-Table DisplayName, Enabled, ProvidedTo, AppId -AutoSize

Write-Host "`n== Who may install own (custom) add-ins? ==" -ForegroundColor Cyan
Get-ManagementRoleAssignment -Role "My Custom Apps" | Format-Table RoleAssigneeName, RoleAssigneeType -AutoSize

Write-Host "`nKlaar. Niets is gewijzigd." -ForegroundColor Green
Disconnect-ExchangeOnline -Confirm:$false
