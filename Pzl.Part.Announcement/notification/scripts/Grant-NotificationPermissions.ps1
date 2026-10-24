<#
.SYNOPSIS
  Grants the application permissions the notification service needs to a
  service principal (user-assigned managed identity or app registration).

.DESCRIPTION
  Assigns these app roles to the service principal:
    Microsoft Graph:            TeamsActivity.Send, GroupMember.Read.All, Sites.Selected
    SharePoint Online:          Sites.Selected
  Requires Microsoft.Graph PowerShell and an account with
  AppRoleAssignment.ReadWrite.All (Global Administrator or Privileged Role Administrator).

  Site-level access (Sites.Selected) is granted separately per site with
  Grant-PnPAzureADAppSitePermission, see README.md.

.PARAMETER ServicePrincipalObjectId
  Object ID of the managed identity's service principal (Azure portal: Managed Identities > Object (principal) ID).

.EXAMPLE
  .\Grant-NotificationPermissions.ps1 -ServicePrincipalObjectId 00000000-0000-0000-0000-000000000000
#>
param(
  [Parameter(Mandatory = $true)]
  [string]$ServicePrincipalObjectId
)

Connect-MgGraph -Scopes 'AppRoleAssignment.ReadWrite.All', 'Application.Read.All' -NoWelcome

$graphAppId = '00000003-0000-0000-c000-000000000000'
$sharePointAppId = '00000003-0000-0ff1-ce00-000000000000'

$required = @(
  @{ Resource = $graphAppId;      Role = 'TeamsActivity.Send' },
  @{ Resource = $graphAppId;      Role = 'GroupMember.Read.All' },
  @{ Resource = $graphAppId;      Role = 'Sites.Selected' },
  @{ Resource = $sharePointAppId; Role = 'Sites.Selected' }
)

foreach ($item in $required) {
  $resource = Get-MgServicePrincipal -Filter "appId eq '$($item.Resource)'"
  $appRole = $resource.AppRoles | Where-Object { $_.Value -eq $item.Role -and $_.AllowedMemberTypes -contains 'Application' }
  if (-not $appRole) { throw "App role $($item.Role) not found on $($resource.DisplayName)" }

  $existing = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ServicePrincipalObjectId |
    Where-Object { $_.AppRoleId -eq $appRole.Id -and $_.ResourceId -eq $resource.Id }
  if ($existing) {
    Write-Host "Already assigned: $($resource.DisplayName) / $($item.Role)"
    continue
  }

  New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ServicePrincipalObjectId -BodyParameter @{
    principalId = $ServicePrincipalObjectId
    resourceId  = $resource.Id
    appRoleId   = $appRole.Id
  } | Out-Null
  Write-Host "Assigned: $($resource.DisplayName) / $($item.Role)"
}
