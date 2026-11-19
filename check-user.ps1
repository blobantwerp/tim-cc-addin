# Read-only check of one mailbox. Changes nothing.
param([Parameter(Mandatory = $true)][string]$Mailbox)

Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline -ShowBanner:$false

Write-Host "`n== Mailbox ==" -ForegroundColor Cyan
Get-Mailbox -Identity $Mailbox | Format-List DisplayName, PrimarySmtpAddress, RecipientTypeDetails, RoleAssignmentPolicy

$policy = (Get-Mailbox -Identity $Mailbox).RoleAssignmentPolicy
Write-Host "`n== Rollen in zijn rolopdrachtenregel ($policy) ==" -ForegroundColor Cyan
Get-ManagementRoleAssignment -RoleAssignee $policy | Where-Object { $_.Role -like "My*" } | Format-Table Role, RoleAssigneeName -AutoSize

Write-Host "`n== Add-ins van deze mailbox ==" -ForegroundColor Cyan
Get-App -Mailbox $Mailbox | Format-Table DisplayName, Enabled, OrgApp, AppId -AutoSize

Write-Host "`n== Client-toegang ==" -ForegroundColor Cyan
Get-CASMailbox -Identity $Mailbox | Format-List OWAEnabled, OutlookMobileEnabled, MAPIEnabled

Write-Host "`nKlaar. Niets is gewijzigd." -ForegroundColor Green
Disconnect-ExchangeOnline -Confirm:$false
