<#
.SYNOPSIS
  Installs or upgrades the Azure, Entra and Teams part of the Teams notification add-on
  (Pzl.Ext.Notification) in one go.

.DESCRIPTION
  Runs the scripts in this folder in order and passes the values between them. Idempotent: run it
  for a new installation and for every upgrade.

    1. Azure resources and app settings (Deploy-NotificationInfrastructure.ps1). The Teams app ID
       and DRY_RUN are read from the existing function app and kept, unless given.
    2. App roles for the managed identity in Entra (Grant-NotificationPermissions.ps1).
    3. Site access for the identity (Sites.Selected): Write on the sites with announcement lists,
       Read on the configuration site and the news-only sites. Existing grants are raised, never
       lowered.
    4. Teams app package (New-NotificationTeamsApp.ps1), uploaded to the organization's app
       catalog with Graph (new app, or a new version of the existing one).
    5. Function code (Publish-NotificationFunction.ps1).

  The SharePoint part (package, list columns, configuration list, deep link custom action) is
  installed with Pzl.Part.Announcement/scripts/Install-Driftsmeldinger.ps1 -EnableTeamsNotifications.
  Left by hand: the Teams app setup policy and the rows in Varslingskonfigurasjon (printed at the end).

  Requires, and does not install: Azure CLI 2.70+ (logged in: az login --tenant <tenant-id>),
  Node 20+ and npm, Microsoft.Graph PowerShell (Authentication and Applications) and PnP.PowerShell.
  Roles: Owner (or Contributor + User Access Administrator) on the subscription, Global
  Administrator or Privileged Role Administrator (app roles), Teams Administrator (app upload)
  and SharePoint Administrator (site access).

.PARAMETER TenantId
  Entra tenant ID.

.PARAMETER SubscriptionId
  Azure subscription for the resources.

.PARAMETER ResourceGroupName
  Resource group (created if missing).

.PARAMETER FunctionAppName
  Function app name (created if missing). The identity is named <FunctionAppName>-identity.

.PARAMETER StorageAccountName
  Storage account name, 3-24 lower case letters and digits (created if missing).

.PARAMETER Location
  Region that supports Flex Consumption (az functionapp list-flexconsumption-locations).

.PARAMETER ConfigSiteUrl
  Site with the Varslingskonfigurasjon list. Also the site the Teams tab opens.

.PARAMETER AnnouncementSiteUrls
  Sites with a Driftsmeldinger list the service notifies from (the identity gets Write, to set
  «Varslet i Teams»). Defaults to ConfigSiteUrl.

.PARAMETER NewsSiteUrls
  Sites the service only reads news from (the identity gets Read).

.PARAMETER AppName
  Name of the Teams app, shown in the bell, e.g. "Svaneposten varsler".

.PARAMETER DeveloperName
  Organization that owns the intranet, shown in the Teams app.

.PARAMETER TeamsAppId
  Teams app ID. Read from the function app setting TEAMS_DEEPLINK_APP_ID when empty, and a new
  GUID is generated for a new installation.

.PARAMETER TeamsAppVersion
  Version of the Teams app package. When empty: 1.0.0 for a new app, otherwise the version in the
  app catalog with the patch number raised.

.PARAMETER ColorIcon
  Color icon for the Teams app (192x192 PNG). Defaults to teams-app/color.png.

.PARAMETER OutlineIcon
  Outline icon for the Teams app (32x32 transparent PNG). Defaults to teams-app/outline.png.

.PARAMETER DryRun
  DRY_RUN app setting: Keep (the current value, On for a new function app), On or Off.

.PARAMETER ClientId
  Client ID used by Connect-PnPOnline -Interactive for the site access step. Defaults to the
  Prosjektportalen (PP365) PnP app.

.PARAMETER SkipInfrastructure
  Skip step 1 (the function app must exist).

.PARAMETER SkipPermissions
  Skip step 2.

.PARAMETER SkipSitePermissions
  Skip step 3.

.PARAMETER SkipTeamsApp
  Skip step 4.

.PARAMETER SkipTeamsAppUpload
  Build the Teams app package in step 4 but do not upload it (upload it in Teams admin center).

.PARAMETER SkipFunction
  Skip step 5.

.EXAMPLE
  az login --tenant <tenant-id>
  .\Install-Notification.ps1 -TenantId <tenant-id> -SubscriptionId <subscription-id> `
    -ResourceGroupName rg-intranet-notification -FunctionAppName func-intranet-notification `
    -StorageAccountName stintranetnotify01 -ConfigSiteUrl https://contoso.sharepoint.com `
    -AppName "Contoso varsler" -DeveloperName "Contoso"

.EXAMPLE
  # Upgrade: new function code and a new Teams app version, everything else as it is
  .\Install-Notification.ps1 -TenantId <tenant-id> -SubscriptionId <subscription-id> `
    -ResourceGroupName rg-intranet-notification -FunctionAppName func-intranet-notification `
    -StorageAccountName stintranetnotify01 -ConfigSiteUrl https://contoso.sharepoint.com `
    -AppName "Contoso varsler" -DeveloperName "Contoso" -SkipSitePermissions
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)] [string]$TenantId,
  [Parameter(Mandatory = $true)] [string]$SubscriptionId,
  [Parameter(Mandatory = $true)] [string]$ResourceGroupName,
  [Parameter(Mandatory = $true)] [string]$FunctionAppName,
  [Parameter(Mandatory = $true)] [ValidatePattern('^[a-z0-9]{3,24}$')] [string]$StorageAccountName,
  [string]$Location = 'westeurope',
  [Parameter(Mandatory = $true)] [string]$ConfigSiteUrl,
  [string[]]$AnnouncementSiteUrls = @(),
  [string[]]$NewsSiteUrls = @(),
  [Parameter(Mandatory = $true)] [string]$AppName,
  [Parameter(Mandatory = $true)] [string]$DeveloperName,
  [string]$TeamsAppId = '',
  [string]$TeamsAppVersion = '',
  [string]$ColorIcon = '',
  [string]$OutlineIcon = '',
  [ValidateSet('Keep', 'On', 'Off')] [string]$DryRun = 'Keep',
  [string]$ClientId = 'da6c31a6-b557-4ac3-9994-7315da06ea3a',
  [switch]$SkipInfrastructure,
  [switch]$SkipPermissions,
  [switch]$SkipSitePermissions,
  [switch]$SkipTeamsApp,
  [switch]$SkipTeamsAppUpload,
  [switch]$SkipFunction
)

$ErrorActionPreference = 'Stop'
$ConfigSiteUrl = $ConfigSiteUrl.TrimEnd('/')
if (-not $AnnouncementSiteUrls.Count) { $AnnouncementSiteUrls = @($ConfigSiteUrl) }
$IdentityName = "$FunctionAppName-identity"
$manualSteps = [System.Collections.Generic.List[string]]::new()

function Write-Step([string]$Text) { Write-Host "`n#### $Text" -ForegroundColor Magenta }
function Write-Done([string]$Text) { Write-Host "   $Text" -ForegroundColor Green }
function Write-Info([string]$Text) { Write-Host "   $Text" }

function Get-AzJson {
  # Runs az and returns parsed JSON, or $null if the command fails (used for lookups)
  param([Parameter(ValueFromRemainingArguments = $true)] [string[]]$Arguments)
  $out = & az @Arguments -o json 2>$null
  if ($LASTEXITCODE -ne 0 -or -not $out) { return $null }
  return (($out -join "`n") | ConvertFrom-Json)
}

#region Prerequisites

Write-Step 'Prerequisites'
$needAz = -not ($SkipInfrastructure -and $SkipFunction)
$needGraph = -not $SkipPermissions -or (-not $SkipTeamsApp -and -not $SkipTeamsAppUpload)
$missing = @()
if (-not (Get-Command az -ErrorAction SilentlyContinue)) { $missing += 'Azure CLI (az)' }
if (-not $SkipFunction -and -not (Get-Command npm -ErrorAction SilentlyContinue)) { $missing += 'Node.js/npm' }
if ($needGraph) {
  foreach ($name in 'Microsoft.Graph.Authentication', 'Microsoft.Graph.Applications') {
    if (-not (Get-Module -ListAvailable -Name $name)) { $missing += "PowerShell module $name" }
  }
}
if (-not $SkipSitePermissions) {
  # Order: already loaded, then the local loader Import-PnPPowerShell, then an installed module
  if (-not (Get-Module -Name PnP.PowerShell) -and (Get-Command Import-PnPPowerShell -ErrorAction SilentlyContinue)) {
    Import-PnPPowerShell | Out-Null
  }
  if (-not (Get-Module -Name PnP.PowerShell)) {
    $available = Get-Module -ListAvailable -Name PnP.PowerShell | Sort-Object Version -Descending | Select-Object -First 1
    if ($available) { Import-Module $available.Path -ErrorAction Stop }
  }
  if (-not (Get-Module -Name PnP.PowerShell)) { $missing += 'PowerShell module PnP.PowerShell (load it in the session or provide Import-PnPPowerShell)' }
}
if ($missing) { throw "Missing (this script installs nothing): $($missing -join ', ')" }

$account = Get-AzJson account show
if (-not $account) { throw "Azure CLI is not logged in. Run: az login --tenant $TenantId" }
if ($account.tenantId -ne $TenantId) { throw "Azure CLI is logged in to tenant $($account.tenantId), expected $TenantId. Run: az login --tenant $TenantId" }
& az account set --subscription $SubscriptionId -o none
if ($LASTEXITCODE -ne 0) { throw "Subscription $SubscriptionId not found for $($account.user.name)" }
Write-Done "Azure CLI: $($account.user.name), subscription $SubscriptionId"

if ($needGraph) {
  $scopes = @('AppRoleAssignment.ReadWrite.All', 'Application.Read.All')
  if (-not $SkipTeamsApp -and -not $SkipTeamsAppUpload) { $scopes += 'AppCatalog.ReadWrite.All' }
  Connect-MgGraph -TenantId $TenantId -Scopes $scopes -NoWelcome
  $context = Get-MgContext
  if ($context.TenantId -ne $TenantId) { throw "Microsoft Graph is connected to tenant $($context.TenantId), expected $TenantId. Run Disconnect-MgGraph and try again." }
  Write-Done "Microsoft Graph: $($context.Account)"
}

#endregion

#region Current state

Write-Step 'Current state'
$functionApp = Get-AzJson functionapp show -n $FunctionAppName -g $ResourceGroupName
$appSettings = @{}
if ($functionApp) {
  $list = Get-AzJson functionapp config appsettings list -n $FunctionAppName -g $ResourceGroupName
  foreach ($setting in @($list)) { $appSettings[$setting.name] = $setting.value }
  Write-Info "Function app $FunctionAppName exists"
} else {
  if ($SkipInfrastructure) { throw "Function app $FunctionAppName not found in $ResourceGroupName; run without -SkipInfrastructure" }
  Write-Info "Function app $FunctionAppName does not exist yet"
}
$mode = if ($functionApp) { 'upgrade' } else { 'new installation' }

if (-not $TeamsAppId) { $TeamsAppId = $appSettings['TEAMS_DEEPLINK_APP_ID'] }
if (-not $TeamsAppId) {
  $TeamsAppId = [guid]::NewGuid().ToString()
  Write-Info "New Teams app ID: $TeamsAppId"
} else {
  Write-Info "Teams app ID: $TeamsAppId"
}

$dryRunValue = switch ($DryRun) {
  'On' { $true }
  'Off' { $false }
  default { if ($appSettings.ContainsKey('DRY_RUN')) { $appSettings['DRY_RUN'] -match '^(true|1|yes)$' } else { $true } }
}
Write-Done "Mode: $mode, DRY_RUN=$($dryRunValue.ToString().ToLower())"

#endregion

#region 1. Azure resources

if ($SkipInfrastructure) {
  Write-Step '1. Azure resources (skipped)'
} else {
  Write-Step '1. Azure resources'
  $deploy = @{
    SubscriptionId     = $SubscriptionId
    ResourceGroupName  = $ResourceGroupName
    Location           = $Location
    FunctionAppName    = $FunctionAppName
    StorageAccountName = $StorageAccountName
    IdentityName       = $IdentityName
    ConfigSiteUrl      = $ConfigSiteUrl
    TeamsDeepLinkAppId = $TeamsAppId
    DryRun             = $dryRunValue
  }
  # Keep settings the infrastructure script would otherwise reset to its defaults
  if ($appSettings['CONFIG_LIST_TITLE']) { $deploy.ConfigListTitle = $appSettings['CONFIG_LIST_TITLE'] }
  if ($appSettings['ANNOUNCEMENT_LIST_TITLE']) { $deploy.AnnouncementListTitle = $appSettings['ANNOUNCEMENT_LIST_TITLE'] }
  if ($appSettings['NEWS_LOOKBACK_DAYS']) { $deploy.NewsLookbackDays = $appSettings['NEWS_LOOKBACK_DAYS'] }
  if ($appSettings['TEAMS_DEEPLINK_ENTITY_ID']) { $deploy.TeamsDeepLinkEntityId = $appSettings['TEAMS_DEEPLINK_ENTITY_ID'] }
  & (Join-Path $PSScriptRoot 'Deploy-NotificationInfrastructure.ps1') @deploy | Out-Host
}

$identity = Get-AzJson identity show -n $IdentityName -g $ResourceGroupName
if (-not $identity) { throw "Managed identity $IdentityName not found in $ResourceGroupName" }
Write-Info "Identity: client ID $($identity.clientId), principal ID $($identity.principalId)"

#endregion

#region 2. App roles

if ($SkipPermissions) {
  Write-Step '2. App roles (skipped)'
} else {
  Write-Step '2. App roles in Entra'
  & (Join-Path $PSScriptRoot 'Grant-NotificationPermissions.ps1') -TenantId $TenantId -ServicePrincipalObjectId $identity.principalId | Out-Host
}

#endregion

#region 3. Site access

if ($SkipSitePermissions) {
  Write-Step '3. Site access (skipped)'
} else {
  Write-Step '3. Site access (Sites.Selected)'
  # Highest level per site: Write wins over Read
  $wanted = [ordered]@{}
  foreach ($url in @($ConfigSiteUrl) + $NewsSiteUrls) { $wanted[$url.TrimEnd('/')] = 'Read' }
  foreach ($url in $AnnouncementSiteUrls) { $wanted[$url.TrimEnd('/')] = 'Write' }
  $rank = @{ read = 1; write = 2; manage = 3; fullcontrol = 4 }

  $connection = Connect-PnPOnline -Url $ConfigSiteUrl -Interactive -ClientId $ClientId -ReturnConnection
  # Connect-PnPOnline does not call SharePoint, and the lookup below ignores errors, so check the
  # sign-in here: a rejected one would otherwise show up as a failed grant.
  try {
    Get-PnPWeb -Connection $connection -ErrorAction Stop | Out-Null
  } catch {
    throw "SharePoint refused the connection to $ConfigSiteUrl ($($_.Exception.Message)). Check that the account belongs to this tenant and has access to the site (sign in again in a private window if the browser picked another work account), that the URL is right, and that the app registration $ClientId is consented for SharePoint in this tenant (otherwise pass -ClientId)."
  }
  foreach ($url in $wanted.Keys) {
    $level = $wanted[$url]
    $grants = @(Get-PnPAzureADAppSitePermission -Site $url -Connection $connection -ErrorAction SilentlyContinue |
        Where-Object { $_.Apps.Id -contains $identity.clientId })
    $current = $grants | Sort-Object { ($_.Roles | ForEach-Object { $rank["$_".ToLower()] } | Measure-Object -Maximum).Maximum } -Descending | Select-Object -First 1
    $currentRank = if ($current) { ($current.Roles | ForEach-Object { $rank["$_".ToLower()] } | Measure-Object -Maximum).Maximum } else { 0 }
    if (-not $current) {
      Grant-PnPAzureADAppSitePermission -AppId $identity.clientId -DisplayName $FunctionAppName -Site $url -Permissions $level -Connection $connection | Out-Null
      Write-Done "${level}: $url (granted)"
    } elseif ($currentRank -lt $rank[$level.ToLower()]) {
      Set-PnPAzureADAppSitePermission -Site $url -PermissionId $current.Id -Permissions $level -Connection $connection | Out-Null
      Write-Done "${level}: $url (raised from $($current.Roles -join ', '))"
    } else {
      Write-Info "$($current.Roles -join ', '): $url (kept)"
    }
  }
}

#endregion

#region 4. Teams app

if ($SkipTeamsApp) {
  Write-Step '4. Teams app (skipped)'
} else {
  Write-Step '4. Teams app'
  $catalogApp = $null
  if (-not $SkipTeamsAppUpload) {
    $found = Invoke-MgGraphRequest -Method GET -Uri "v1.0/appCatalogs/teamsApps?`$filter=externalId eq '$TeamsAppId'&`$expand=appDefinitions"
    $catalogApp = @($found.value) | Select-Object -First 1
  }
  if (-not $TeamsAppVersion) {
    if ($catalogApp) {
      $versions = @($catalogApp.appDefinitions | ForEach-Object { [version]$_.version })
      $latest = $versions | Sort-Object -Descending | Select-Object -First 1
      $TeamsAppVersion = "$($latest.Major).$($latest.Minor).$([math]::Max($latest.Build, 0) + 1)"
    } else {
      $TeamsAppVersion = '1.0.0'
    }
  }
  Write-Info "Version $TeamsAppVersion"

  $output = Join-Path ([System.IO.Path]::GetTempPath()) "teamsapp-out-$([guid]::NewGuid())"
  $build = @{
    NotificationClientId = $identity.clientId
    SiteUrl              = $ConfigSiteUrl
    AppName              = $AppName
    DeveloperName        = $DeveloperName
    AppId                = $TeamsAppId
    Version              = $TeamsAppVersion
    OutputPath           = $output
  }
  if ($ColorIcon) { $build.ColorIcon = $ColorIcon }
  if ($OutlineIcon) { $build.OutlineIcon = $OutlineIcon }
  & (Join-Path $PSScriptRoot 'New-NotificationTeamsApp.ps1') @build | Out-Host
  $zip = Get-ChildItem $output -Filter '*.zip' | Select-Object -First 1
  if (-not $zip) { throw "Teams app package not found in $output" }

  if ($SkipTeamsAppUpload) {
    Write-Done "Package: $($zip.FullName)"
    $manualSteps.Add("Upload $($zip.FullName) in Teams admin center (Teams apps > Manage apps).")
  } elseif ($catalogApp) {
    Invoke-MgGraphRequest -Method POST -Uri "v1.0/appCatalogs/teamsApps/$($catalogApp.id)/appDefinitions" -InputFilePath $zip.FullName -ContentType 'application/zip' | Out-Null
    Write-Done "Uploaded version $TeamsAppVersion of $AppName to the app catalog"
    $manualSteps.Add('Teams clients pick up the new app version within hours; sign out and in of Teams to test it at once.')
  } else {
    Invoke-MgGraphRequest -Method POST -Uri 'v1.0/appCatalogs/teamsApps' -InputFilePath $zip.FullName -ContentType 'application/zip' | Out-Null
    Write-Done "Published $AppName ($TeamsAppId) to the organization's app catalog"
    $manualSteps.Add("Teams admin center: create an app setup policy that installs «$AppName» for the recipients (without pinning it) and assign it to the same groups as in Varslingskonfigurasjon. Start with a test group.")
  }
}

#endregion

#region 5. Function code

if ($SkipFunction) {
  Write-Step '5. Function code (skipped)'
} else {
  Write-Step '5. Function code'
  & (Join-Path $PSScriptRoot 'Publish-NotificationFunction.ps1') -ResourceGroupName $ResourceGroupName -FunctionAppName $FunctionAppName | Out-Host
}

#endregion

#region Summary

Write-Step 'Done'
Write-Done "Teams notification add-on ($mode): $FunctionAppName, Teams app $TeamsAppId"
$manualSteps.Insert(0, "SharePoint part, if not done: Pzl.Part.Announcement/scripts/Install-Driftsmeldinger.ps1 -SiteUrl <list site> -EnableTeamsNotifications -NotificationSiteUrl $ConfigSiteUrl")
$manualSteps.Add("Add one row per site to Varslingskonfigurasjon on $ConfigSiteUrl (see README, step 3).")
$manualSteps.Add("Send a test notification to one user: README, step 6 (/api/test-notification).")
if ($dryRunValue) {
  $manualSteps.Add("DRY_RUN is on. Check the log with .\Get-NotificationLog.ps1 -ResourceGroupName $ResourceGroupName -FunctionAppName $FunctionAppName, then run again with -DryRun Off.")
}
Write-Host "`nManual steps:" -ForegroundColor Yellow
$manualSteps | ForEach-Object { Write-Host "  - $_" }

#endregion
