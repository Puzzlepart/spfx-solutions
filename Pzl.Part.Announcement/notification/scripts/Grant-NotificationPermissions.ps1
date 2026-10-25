<#
.SYNOPSIS
  Grants the application permissions the notification service needs to a
  service principal (user-assigned managed identity or app registration).

.DESCRIPTION
  Assigns these app roles to the service principal:
    Microsoft Graph:            TeamsActivity.Send, GroupMember.Read.All, Sites.Selected
    SharePoint Online:          Sites.Selected
  Requires Microsoft.Graph PowerShell and an account with
  AppRoleAssignment.ReadWrite.All (Global Administrator or Privileged Role Administrator)
  in the tenant that holds the identity.

  Site-level access (Sites.Selected) is granted separately per site with
  Grant-PnPAzureADAppSitePermission, see README.md.

.PARAMETER TenantId
  Tenant that holds the identity. Required, because Connect-MgGraph otherwise reuses the last
  tenant you signed in to, and the identity will then not be found.

.PARAMETER ServicePrincipalObjectId
  Object ID of the managed identity's service principal (printed by
  Deploy-NotificationInfrastructure.ps1 as IdentityPrincipalId; Azure portal: Managed Identities > Object (principal) ID).

.EXAMPLE
  .\Grant-NotificationPermissions.ps1 -TenantId 00000000-0000-0000-0000-000000000000 -ServicePrincipalObjectId 11111111-1111-1111-1111-111111111111
#>
param(
  [Parameter(Mandatory = $true)]
  [string]$TenantId,
  [Parameter(Mandatory = $true)]
  [string]$ServicePrincipalObjectId
)

$ErrorActionPreference = 'Stop'

Connect-MgGraph -TenantId $TenantId -Scopes 'AppRoleAssignment.ReadWrite.All', 'Application.Read.All' -NoWelcome
$context = Get-MgContext
if ($context.TenantId -ne $TenantId) { throw "Connected to tenant $($context.TenantId), expected $TenantId. Run Disconnect-MgGraph and try again." }
Write-Host "Connected to tenant $($context.TenantId) as $($context.Account)"

try {
  $principal = Get-MgServicePrincipal -ServicePrincipalId $ServicePrincipalObjectId
} catch {
  throw "Service principal $ServicePrincipalObjectId was not found in tenant $TenantId. Use the object (principal) ID of the managed identity, not its client ID, and check the tenant."
}
Write-Host "Service principal: $($principal.DisplayName) (appId $($principal.AppId), type $($principal.ServicePrincipalType))"

$graphAppId = '00000003-0000-0000-c000-000000000000'
$sharePointAppId = '00000003-0000-0ff1-ce00-000000000000'

$required = @(
  @{ Resource = $graphAppId;      Role = 'TeamsActivity.Send' },
  @{ Resource = $graphAppId;      Role = 'GroupMember.Read.All' },
  @{ Resource = $graphAppId;      Role = 'Sites.Selected' },
  @{ Resource = $sharePointAppId; Role = 'Sites.Selected' }
)

$existingAssignments = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ServicePrincipalObjectId -All

foreach ($item in $required) {
  $resource = Get-MgServicePrincipal -Filter "appId eq '$($item.Resource)'"
  if (-not $resource) { throw "Resource service principal $($item.Resource) not found in the tenant" }
  $appRole = $resource.AppRoles | Where-Object { $_.Value -eq $item.Role -and $_.AllowedMemberTypes -contains 'Application' }
  if (-not $appRole) { throw "App role $($item.Role) not found on $($resource.DisplayName)" }

  $existing = $existingAssignments | Where-Object { $_.AppRoleId -eq $appRole.Id -and $_.ResourceId -eq $resource.Id }
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

Write-Host ''
Write-Host 'Done. App roles take effect for new tokens within a few minutes.' -ForegroundColor Green
Write-Host 'Next: grant site access per site, e.g.'
Write-Host "  Grant-PnPAzureADAppSitePermission -AppId $($principal.AppId) -DisplayName '$($principal.DisplayName)' -Site https://<tenant>.sharepoint.com -Permissions Write"
