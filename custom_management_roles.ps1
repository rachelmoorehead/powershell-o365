## Setup Custom Management Role
# 
# Turn on Customizations, once on, always on
Enable-OrganizationCustomization

# Create new Management Role based on original
New-ManagementRole -Name Custom_MyBaseOptions -Parent MyBaseOptions
# See what cmdlets it contains
Get-ManagementRole MyBaseOptions| select -expand RoleEntries | Out-File C:\PowerShell\mybaseoptions_roles.txt

# Modify parameters that can be used, in this case password resets
Set-ManagementRoleEntry Custom_MyBaseOptions\Set-Mailbox -Parameters Password -RemoveParameter
Set-ManagementRoleEntry Custom_MyBaseOptions\Set-MailUser -Parameters Password -RemoveParameter

# See the original assignment to get name
Get-ManagementRoleAssignment | Where-Object {$_.Name -match "MyBaseOptions"} | select Name
# Create a new assignment using the same naming convention prepended with Custom
New-ManagementRoleAssignment -Name "Custom_MyBaseOptions-Default Role Assignment Policy" -Role Custom_MyBaseOptions -Policy "Default Role Assignment Policy"
# Verify creation
Get-ManagementRoleAssignment | Where-Object { $_.RoleAssigneeName -match "Default Role Assignment Policy" }
# Remove default assignment, not role
Remove-ManagementRoleAssignment "MyBaseOptions-Default Role Assignment Policy"

# Repeat with other Roles such as Distribution Groups
New-ManagementRole -Name Custom_MyDistributionGroups -Parent MyDistributionGroups
Get-ManagementRole MyDistributionGroups| select -expand RoleEntries | Out-File C:\PowerShell\mydistributiongroups_roles.txt

Remove-ManagementRoleEntry Custom_MyDistributionGroups\New-DistributionGroup
Remove-ManagementRoleEntry Custom_MyDistributionGroups\Remove-DistributionGroup

Set-ManagementRoleEntry Custom_MyDistributionGroups\Set-DistributionGroup -Parameters EmailAddresses,PrimarySmtpAddress,WindowsEmailAddress -RemoveParameter
Set-ManagementRoleEntry Custom_MyDistributionGroups\Set-Group -Parameters WindowsEmailAddress -RemoveParameter

Get-ManagementRoleAssignment | Where-Object { $_.Name -match "MyDistributionGroups" } | select Name
New-ManagementRoleAssignment -Name "Custom_MyDistributionGroups-Default Role Assignment Policy" -Role Custom_MyDistributionGroups -Policy "Default Role Assignment Policy"

Get-ManagementRoleAssignment | Where-Object { $_.RoleAssigneeName -match "Default Role Assignment Policy" }
Remove-ManagementRoleAssignment "MyDistributionGroups-Default Role Assignment Policy"

#Note: Instantaneous for new accounts but not for existing.  Could take up to one hour to propagate to all accounts changed.
#