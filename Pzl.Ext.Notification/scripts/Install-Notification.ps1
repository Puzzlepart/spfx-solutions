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

  Requires, and does not install: Azure CLI 2.70+ (the script runs az login --tenant <tenant-id>
  when the CLI is not logged in to the tenant, except with -Force),
  Microsoft.Graph PowerShell (Authentication and Applications) and PnP.PowerShell. Node 20+ and npm
  are needed only when the function is built from source: a release package holds the function
  prebuilt (function/notification-function.zip, no function/src), and that file is deployed as it is.
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
  Region for the Azure resources, default Norway East. Must support Flex Consumption
  (az functionapp list-flexconsumption-locations), which is checked before anything is created.

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
  Manifest ID of the notification Teams app (the app this script builds and uploads). When empty
  it is read from the function app setting TEAMS_APP_MANIFEST_ID, which this script writes. For
  installations made before that setting existed it is read from TEAMS_DEEPLINK_APP_ID, but only
  when the notifications open the notification app's own tab. A new GUID is generated for a new
  installation.

.PARAMETER TeamsDeepLinkAppId
  Teams app the notifications open (TEAMS_DEEPLINK_APP_ID). Pass the ID of the intranet's own
  Teams app to open its tab; -TeamsDeepLinkEntityId is then required. When empty the current
  setting is kept, and a new installation opens the notification app's own tab.

.PARAMETER TeamsDeepLinkEntityId
  entityId of the tab in that app (TEAMS_DEEPLINK_ENTITY_ID). "intranet" is the notification
  app's own tab.

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

.PARAMETER PreflightOnly
  Show the preflight summary (environment, what is created or changed, warnings) and stop.
  Nothing is created or changed.

.PARAMETER GraphFromAzureCli
  Use the sign-in of Azure CLI for Microsoft Graph instead of a sign-in window. Use it when the
  window (WAM, which the Graph module cannot turn off) does not show up and the script hangs,
  typically in a terminal inside another application. The account must be an administrator who
  may assign app roles. If the sign-in has no access to the Teams app catalog, the package is
  built and uploaded by hand in Teams admin center.

.PARAMETER GraphDeviceCode
  Sign in to Microsoft Graph with a device code (a code to type in at microsoft.com/devicelogin)
  instead of a sign-in window. Use it when the window does not show up and the script hangs. A
  session already signed in to the tenant with the scopes needed is used without a new sign-in.

.PARAMETER Force
  Do not ask for confirmation after the preflight summary (for unattended runs). The script
  then does not sign in to Azure CLI itself: it stops if the CLI is not logged in to the tenant.

.EXAMPLE
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

.EXAMPLE
  # The notifications open the tab of the intranet's own Teams app
  .\Install-Notification.ps1 -TenantId <tenant-id> -SubscriptionId <subscription-id> `
    -ResourceGroupName rg-intranet-notification -FunctionAppName func-intranet-notification `
    -StorageAccountName stintranetnotify01 -ConfigSiteUrl https://contoso.sharepoint.com `
    -AppName "Contoso varsler" -DeveloperName "Contoso" `
    -TeamsDeepLinkAppId <intranet app ID> -TeamsDeepLinkEntityId <entityId of its tab>
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)] [string]$TenantId,
  [Parameter(Mandatory = $true)] [string]$SubscriptionId,
  [Parameter(Mandatory = $true)] [string]$ResourceGroupName,
  [Parameter(Mandatory = $true)] [string]$FunctionAppName,
  [Parameter(Mandatory = $true)] [ValidatePattern('^[a-z0-9]{3,24}$')] [string]$StorageAccountName,
  [string]$Location = 'norwayeast',
  [Parameter(Mandatory = $true)] [string]$ConfigSiteUrl,
  [string[]]$AnnouncementSiteUrls = @(),
  [string[]]$NewsSiteUrls = @(),
  [Parameter(Mandatory = $true)] [string]$AppName,
  [Parameter(Mandatory = $true)] [string]$DeveloperName,
  [string]$TeamsAppId = '',
  [string]$TeamsAppVersion = '',
  [string]$TeamsDeepLinkAppId = '',
  [string]$TeamsDeepLinkEntityId = '',
  [string]$ColorIcon = '',
  [string]$OutlineIcon = '',
  [ValidateSet('Keep', 'On', 'Off')] [string]$DryRun = 'Keep',
  [string]$ClientId = 'da6c31a6-b557-4ac3-9994-7315da06ea3a',
  [switch]$SkipInfrastructure,
  [switch]$SkipPermissions,
  [switch]$SkipSitePermissions,
  [switch]$SkipTeamsApp,
  [switch]$SkipTeamsAppUpload,
  [switch]$SkipFunction,
  [switch]$PreflightOnly,
  [switch]$GraphDeviceCode,
  [switch]$GraphFromAzureCli,
  [switch]$Force
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

function Find-CatalogApp([string]$ExternalId) {
  # Finds a Teams app in the organization's app catalog by its manifest ID. The externalId filter
  # of the API has answered with nothing for an app that was in the catalog, so the organization's
  # apps are listed and matched here instead.
  $uri = "v1.0/appCatalogs/teamsApps?`$filter=distributionMethod eq 'organization'&`$expand=appDefinitions"
  while ($uri) {
    $page = Invoke-MgGraphRequest -Method GET -Uri $uri
    $match = @($page.value) | Where-Object { $_.externalId -eq $ExternalId } | Select-Object -First 1
    if ($match) { return $match }
    $uri = $page.'@odata.nextLink'
  }
}

#region Prerequisites

Write-Step 'Prerequisites'
$needAz = -not ($SkipInfrastructure -and $SkipFunction)
# The preflight reads nothing from Microsoft Graph, so -PreflightOnly does not sign in to it
$needGraph = -not $PreflightOnly -and (-not $SkipPermissions -or (-not $SkipTeamsApp -and -not $SkipTeamsAppUpload))
$missing = @()
if (-not (Get-Command az -ErrorAction SilentlyContinue)) { $missing += 'Azure CLI (az)' }
# A release package holds the function prebuilt and no source: nothing to build, so no Node
$functionFolder = Join-Path $PSScriptRoot '..\function'
$prebuiltFunction = (Test-Path (Join-Path $functionFolder 'notification-function.zip')) -and -not (Test-Path (Join-Path $functionFolder 'src'))
if (-not $SkipFunction -and -not $prebuiltFunction -and -not (Get-Command npm -ErrorAction SilentlyContinue)) { $missing += 'Node.js/npm' }
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
if (-not $account -or $account.tenantId -ne $TenantId) {
  $reason = if ($account) { "Azure CLI is logged in to tenant $($account.tenantId), expected $TenantId." } else { 'Azure CLI is not logged in.' }
  # An unattended run cannot answer a sign-in in the browser
  if ($Force) { throw "$reason Run: az login --tenant $TenantId" }
  Write-Info "$reason Signing in (az login --tenant $TenantId)"
  # Without the subscription picker of the new login experience, which waits for an answer
  $loginExperience = $env:AZURE_CORE_LOGIN_EXPERIENCE_V2
  $env:AZURE_CORE_LOGIN_EXPERIENCE_V2 = 'off'
  try {
    & az login --tenant $TenantId -o none
  } finally {
    $env:AZURE_CORE_LOGIN_EXPERIENCE_V2 = $loginExperience
  }
  if ($LASTEXITCODE -ne 0) { throw "az login failed. Run: az login --tenant $TenantId" }
  $account = Get-AzJson account show
  if (-not $account -or $account.tenantId -ne $TenantId) { throw "Azure CLI is still not logged in to tenant $TenantId. Run: az login --tenant $TenantId" }
}
& az account set --subscription $SubscriptionId -o none
if ($LASTEXITCODE -ne 0) { throw "Subscription $SubscriptionId not found for $($account.user.name)" }
Write-Done "Azure CLI: $($account.user.name), subscription $SubscriptionId"

if ($needGraph) {
  $scopes = @('AppRoleAssignment.ReadWrite.All', 'Application.Read.All')
  if (-not $SkipTeamsApp -and -not $SkipTeamsAppUpload) { $scopes += 'AppCatalog.ReadWrite.All' }
  # A session that is already signed in to the tenant with the scopes is used as it is. The
  # sign-in window of Connect-MgGraph has hung without showing and without answering Ctrl+C, so
  # this lets the person sign in first (Connect-MgGraph ... -UseDeviceCode) and run the script after.
  $context = Get-MgContext
  $missingScopes = @($scopes | Where-Object { @($context.Scopes) -notcontains $_ })
  if ($GraphFromAzureCli) {
    # Azure CLI is signed in already, and its Microsoft Graph token lets an administrator assign
    # app roles. It may not cover the Teams app catalog, which is checked below.
    Write-Info 'Microsoft Graph: using the sign-in of Azure CLI'
    $graphToken = & az account get-access-token --resource-type ms-graph --tenant $TenantId --query accessToken -o tsv
    if ($LASTEXITCODE -ne 0 -or -not $graphToken) { throw 'Azure CLI did not return a Microsoft Graph token (az account get-access-token --resource-type ms-graph).' }
    Connect-MgGraph -AccessToken (ConvertTo-SecureString "$graphToken" -AsPlainText -Force) -NoWelcome
    if (-not $SkipTeamsApp -and -not $SkipTeamsAppUpload) {
      try {
        Invoke-MgGraphRequest -Method GET -Uri "v1.0/appCatalogs/teamsApps?`$top=1" | Out-Null
      } catch {
        Write-Warning "The Azure CLI sign-in has no access to the Teams app catalog, so the Teams app package is built but not uploaded: upload it in Teams admin center (printed at the end). For an upgrade, pass -TeamsAppVersion with a higher version than the one in the catalog."
        $SkipTeamsAppUpload = $true
      }
    }
  } elseif ($context -and $context.TenantId -eq $TenantId -and -not $missingScopes.Count) {
    Write-Info 'Microsoft Graph: using the sign-in of this session'
  } elseif ($GraphDeviceCode) {
    Write-Info "Signing in to Microsoft Graph with a device code ($($scopes -join ', '))"
    Connect-MgGraph -TenantId $TenantId -Scopes $scopes -NoWelcome -UseDeviceCode
  } else {
    Write-Info "Signing in to Microsoft Graph ($($scopes -join ', ')). A sign-in window or browser tab opens, and it can be hidden behind other windows. If nothing shows up, close this terminal and run again with -GraphFromAzureCli."
    Connect-MgGraph -TenantId $TenantId -Scopes $scopes -NoWelcome
  }
  $context = Get-MgContext
  # A session made from an access token may not say which tenant it is for
  if ($context.TenantId -and $context.TenantId -ne $TenantId) { throw "Microsoft Graph is connected to tenant $($context.TenantId), expected $TenantId. Run Disconnect-MgGraph and try again." }
  Write-Done "Microsoft Graph: $(if ($context.Account) { $context.Account } else { 'access token from Azure CLI' })"
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

# The notification app's own ID and the app the notifications open are two things. They are the
# same only when the notifications open the notification app's own tab (entity "intranet"). With
# an intranet app as the target, TEAMS_DEEPLINK_APP_ID is that app's ID, and reading the
# notification app's ID from it would upload the notification manifest over the intranet app.
$OwnTabEntityId = 'intranet'
$currentDeepLinkAppId = "$($appSettings['TEAMS_DEEPLINK_APP_ID'])".Trim()
$currentEntityId = "$($appSettings['TEAMS_DEEPLINK_ENTITY_ID'])".Trim()
$currentOpensOwnTab = -not $currentEntityId -or $currentEntityId -eq $OwnTabEntityId

if (-not $TeamsAppId) { $TeamsAppId = "$($appSettings['TEAMS_APP_MANIFEST_ID'])".Trim() }
if (-not $TeamsAppId -and $currentDeepLinkAppId) {
  if (-not $currentOpensOwnTab) {
    throw "The notifications open the tab '$currentEntityId' of the Teams app $currentDeepLinkAppId, which is not the notification app, so the notification app's ID cannot be read from TEAMS_DEEPLINK_APP_ID. Pass -TeamsAppId with the app ID of «$AppName» (Teams admin center > Teams apps > Manage apps > the app > App ID). If the app was never uploaded (an earlier run stopped before step 4), pass the ID that run printed, or a new GUID. The script stores it for later runs."
  }
  $TeamsAppId = $currentDeepLinkAppId
}
if (-not $TeamsAppId) {
  $TeamsAppId = [guid]::NewGuid().ToString()
  Write-Info "New Teams app ID: $TeamsAppId"
} else {
  Write-Info "Teams app ID: $TeamsAppId"
}

# Where the notifications open: as given, otherwise as it is, otherwise the notification app's tab
if ($TeamsDeepLinkAppId) {
  $deepLinkAppId = $TeamsDeepLinkAppId.Trim()
  if ($TeamsDeepLinkEntityId) {
    $deepLinkEntityId = $TeamsDeepLinkEntityId.Trim()
  } elseif ($deepLinkAppId -eq $TeamsAppId) {
    $deepLinkEntityId = $OwnTabEntityId
  } elseif ($deepLinkAppId -eq $currentDeepLinkAppId -and -not $currentOpensOwnTab) {
    $deepLinkEntityId = $currentEntityId
  } else {
    throw "-TeamsDeepLinkAppId $deepLinkAppId is another app than the notification app: pass -TeamsDeepLinkEntityId with the entityId of the tab to open (staticTabs in that app's manifest)."
  }
} elseif ($currentDeepLinkAppId -and -not $currentOpensOwnTab) {
  $deepLinkAppId = $currentDeepLinkAppId
  $deepLinkEntityId = if ($TeamsDeepLinkEntityId) { $TeamsDeepLinkEntityId.Trim() } else { $currentEntityId }
} else {
  $deepLinkAppId = $TeamsAppId
  $deepLinkEntityId = if ($TeamsDeepLinkEntityId) { $TeamsDeepLinkEntityId.Trim() } else { $OwnTabEntityId }
}
$opensOtherApp = $deepLinkAppId -ne $TeamsAppId
if ($opensOtherApp) {
  Write-Info "Notifications open another Teams app: $deepLinkAppId, tab $deepLinkEntityId"
} else {
  Write-Info "Notifications open the notification app's own tab ($deepLinkEntityId)"
}

$dryRunValue = switch ($DryRun) {
  'On' { $true }
  'Off' { $false }
  default { if ($appSettings.ContainsKey('DRY_RUN')) { $appSettings['DRY_RUN'] -match '^(true|1|yes)$' } else { $true } }
}
Write-Done "Mode: $mode, DRY_RUN=$($dryRunValue.ToString().ToLower())"

#endregion

#region Preflight

# Everything above only reads. Show where the script is about to work and what it will do, check
# what can be checked up front, and let the person running it decide.
Write-Step 'Preflight'
$warnings = [System.Collections.Generic.List[string]]::new()
function Write-Row([string]$Label, [string]$Value) { Write-Host ("   {0,-22} {1}" -f $Label, $Value) }
function Write-Section([string]$Text) { Write-Host "`n   $Text" -ForegroundColor Cyan }

# Highest level per site: Write wins over Read
$siteAccess = [ordered]@{}
foreach ($url in @($ConfigSiteUrl) + $NewsSiteUrls) { $siteAccess[$url.TrimEnd('/')] = 'Read' }
foreach ($url in $AnnouncementSiteUrls) { $siteAccess[$url.TrimEnd('/')] = 'Write' }

$subscription = Get-AzJson account show
Write-Section 'Environment'
Write-Row 'Tenant' $TenantId
Write-Row 'Azure account' "$($subscription.user.name)"
Write-Row 'Subscription' "$($subscription.name) ($SubscriptionId)"
if ($needGraph) { Write-Row 'Microsoft Graph' "$((Get-MgContext).Account)" } elseif ($PreflightOnly) { Write-Row 'Microsoft Graph' 'not signed in (-PreflightOnly)' }
Write-Row 'Intranet site' $ConfigSiteUrl

if ($SkipInfrastructure) {
  Write-Section '1. Azure resources: skipped'
} else {
  $group = Get-AzJson group show -n $ResourceGroupName
  $identityNow = Get-AzJson identity show -n $IdentityName -g $ResourceGroupName
  $storageNow = Get-AzJson storage account show -n $StorageAccountName -g $ResourceGroupName
  $state = { param($Existing) if ($Existing) { 'exists, settings updated' } else { "CREATED in $Location" } }
  Write-Section "1. Azure resources ($mode)"
  Write-Row 'Resource group' "$ResourceGroupName ($(if ($group) { "exists, $($group.location)" } else { "CREATED in $Location" }))"
  Write-Row 'Managed identity' "$IdentityName ($(if ($identityNow) { 'exists' } else { "CREATED in $Location" }))"
  Write-Row 'Storage account' "$StorageAccountName ($(if ($storageNow) { 'exists' } else { "CREATED in $Location" }))"
  Write-Row 'Application Insights' "$FunctionAppName-insights (created if missing)"
  Write-Row 'Function app' "$FunctionAppName ($(& $state $functionApp))"
  Write-Row 'DRY_RUN' "$($dryRunValue.ToString().ToLower())$(if ($dryRunValue) { ' (nothing is sent)' } else { ' (NOTIFICATIONS ARE SENT)' })"

  if (-not $storageNow) {
    $nameCheck = Get-AzJson storage account check-name --name $StorageAccountName
    if ($nameCheck -and -not $nameCheck.nameAvailable) {
      $warnings.Add("The storage account name $StorageAccountName is not available ($($nameCheck.message)). Pass another -StorageAccountName.")
    }
  }

  # Resource providers are registered per subscription. az registers a missing one on first use,
  # but only for an account with rights on the subscription: with rights on the resource group
  # alone the creation fails halfway. Checked only when something is to be created.
  if (-not ($group -and $identityNow -and $storageNow -and $functionApp)) {
    $providers = [ordered]@{
      'Microsoft.ManagedIdentity'     = 'managed identity'
      'Microsoft.Storage'             = 'storage account'
      'Microsoft.Insights'            = 'Application Insights'
      'Microsoft.OperationalInsights' = 'Log Analytics workspace for Application Insights'
      'Microsoft.Web'                 = 'function app'
      'Microsoft.App'                 = 'Flex Consumption plan'
    }
    $notRegistered = @()
    foreach ($namespace in $providers.Keys) {
      $provider = Get-AzJson provider show -n $namespace
      $registration = if ($provider) { "$($provider.registrationState)" } else { 'unknown' }
      if ($registration -ne 'Registered') { $notRegistered += "$namespace ($($providers[$namespace]): $registration)" }
    }
    if ($notRegistered) {
      Write-Row 'Resource providers' "NOT REGISTERED: $($notRegistered -join ', ')"
      $warnings.Add("Resource providers not registered in the subscription: $($notRegistered -join ', '). Azure CLI registers them on first use if the account has rights on the subscription. With rights on the resource group only, ask a subscription owner to run: az provider register -n <namespace> --subscription $SubscriptionId")
    } else {
      Write-Row 'Resource providers' "registered ($($providers.Keys.Count) checked)"
    }
  }
  if (-not $group) {
    $warnings.Add("The resource group $ResourceGroupName does not exist and is created, which takes rights on the subscription.")
  }
}

if ($SkipPermissions) {
  Write-Section '2. App roles in Entra: skipped'
} else {
  Write-Section '2. App roles in Entra (for the managed identity)'
  Write-Row 'Microsoft Graph' 'TeamsActivity.Send, GroupMember.Read.All, Sites.Selected'
  Write-Row 'SharePoint Online' 'Sites.Selected'
}

if ($SkipSitePermissions) {
  Write-Section '3. Site access: skipped'
} else {
  Write-Section "3. Site access for the identity (signs in with PnP app $ClientId)"
  foreach ($url in $siteAccess.Keys) { Write-Row $siteAccess[$url] $url }
  Write-Row '' 'Existing grants are raised, never lowered.'
  Write-Row '' 'On a site that refuses the grant, the signed-in account is made site collection'
  Write-Row '' 'administrator for the grant and removed afterwards.'
}

if ($SkipTeamsApp) {
  Write-Section '4. Teams app: skipped'
} else {
  Write-Section '4. Teams app'
  Write-Row 'Name' "$AppName ($DeveloperName)"
  Write-Row 'App ID' $TeamsAppId
  Write-Row 'Icons' "$(if ($ColorIcon) { $ColorIcon } else { 'default color icon' }), $(if ($OutlineIcon) { $OutlineIcon } else { 'default outline icon' })"
  Write-Row 'App catalog' "$(if ($SkipTeamsAppUpload) { 'package is built, NOT uploaded' } else { 'uploaded to the organization''s app catalog (new app, or a new version)' })"
  foreach ($icon in $ColorIcon, $OutlineIcon) {
    if ($icon -and -not (Test-Path -LiteralPath $icon)) { $warnings.Add("Icon not found: $icon") }
  }
}
Write-Row 'Notifications open' "$(if ($opensOtherApp) { "another Teams app: $deepLinkAppId, tab $deepLinkEntityId" } else { "the notification app's own tab ($deepLinkEntityId)" })"

Write-Section "5. Function code: $(if ($SkipFunction) { 'skipped' } elseif ($prebuiltFunction) { "prebuilt package deployed to $FunctionAppName" } else { "built and deployed to $FunctionAppName" })"

if ($warnings.Count) {
  Write-Host ''
  foreach ($warning in $warnings) { Write-Warning $warning }
}
if ($PreflightOnly) {
  Write-Host "`n   -PreflightOnly: nothing was created or changed." -ForegroundColor Yellow
  return
}
if (-not $Force) {
  Write-Host ''
  $question = if ($warnings.Count) { "Continue in spite of $($warnings.Count) warning(s)? (y/N)" } else { 'Continue? (y/N)' }
  try {
    $answer = Read-Host "   $question"
  } catch {
    throw "The preflight needs a confirmation and the session is not interactive. Pass -Force to continue without asking, or -PreflightOnly to see the summary only."
  }
  if ($answer -notmatch '^\s*(y|yes|j|ja)\s*$') {
    Write-Host '   Stopped. Nothing was created or changed.' -ForegroundColor Yellow
    return
  }
}

#endregion

#region 1. Azure resources

if ($SkipInfrastructure) {
  Write-Step '1. Azure resources (skipped)'
  if ($deepLinkAppId -ne $currentDeepLinkAppId -or ($currentDeepLinkAppId -and $deepLinkEntityId -ne $currentEntityId -and -not ($currentOpensOwnTab -and $deepLinkEntityId -eq $OwnTabEntityId))) {
    Write-Warning "TEAMS_DEEPLINK_APP_ID and TEAMS_DEEPLINK_ENTITY_ID are set in step 1, which is skipped: the notifications still open $currentDeepLinkAppId."
  }
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
    TeamsDeepLinkAppId    = $deepLinkAppId
    TeamsDeepLinkEntityId = $deepLinkEntityId
    DryRun                = $dryRunValue
  }
  # Keep settings the infrastructure script would otherwise reset to its defaults
  if ($appSettings['CONFIG_LIST_TITLE']) { $deploy.ConfigListTitle = $appSettings['CONFIG_LIST_TITLE'] }
  if ($appSettings['ANNOUNCEMENT_LIST_TITLE']) { $deploy.AnnouncementListTitle = $appSettings['ANNOUNCEMENT_LIST_TITLE'] }
  if ($appSettings['NEWS_LOOKBACK_DAYS']) { $deploy.NewsLookbackDays = $appSettings['NEWS_LOOKBACK_DAYS'] }
  & (Join-Path $PSScriptRoot 'Deploy-NotificationInfrastructure.ps1') @deploy | Out-Host
}

$identity = Get-AzJson identity show -n $IdentityName -g $ResourceGroupName
if (-not $identity) { throw "Managed identity $IdentityName not found in $ResourceGroupName" }
Write-Info "Identity: client ID $($identity.clientId), principal ID $($identity.principalId)"

# Remember the notification app's ID as soon as the function app exists, so a run that stops in a
# later step can be run again without -TeamsAppId. (The ID is stored again after step 4 if a
# different one was given.)
if (-not "$($appSettings['TEAMS_APP_MANIFEST_ID'])") {
  & az functionapp config appsettings set -n $FunctionAppName -g $ResourceGroupName --settings "TEAMS_APP_MANIFEST_ID=$TeamsAppId" -o none
  if ($LASTEXITCODE -ne 0) { throw "Could not store TEAMS_APP_MANIFEST_ID on $FunctionAppName" }
  $appSettings['TEAMS_APP_MANIFEST_ID'] = $TeamsAppId
  Write-Info "Stored the notification app's ID ($TeamsAppId) in the app setting TEAMS_APP_MANIFEST_ID"
}

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
  $wanted = $siteAccess
  $rank = @{ read = 1; write = 2; manage = 3; fullcontrol = 4 }

  # Without -Tenant, PnP.PowerShell signs in against the common endpoint and silently reuses the
  # first account it has cached for the client ID (a persisted login included), also one from
  # another tenant, and SharePoint answers such a token with "Unauthorized". Connect-PnPOnline does
  # not call SharePoint and the lookup below ignores errors, so read the web here, and ask for a
  # new sign-in once if the token is refused.
  $connection = $null
  $failure = ''
  foreach ($force in $false, $true) {
    if ($force) { Write-Warning "SharePoint refused the sign-in to $ConfigSiteUrl ($failure). Asking for a new sign-in." }
    $connection = Connect-PnPOnline -Url $ConfigSiteUrl -Interactive -ClientId $ClientId -Tenant $TenantId -ForceAuthentication:$force -ReturnConnection
    try {
      Get-PnPWeb -Connection $connection -ErrorAction Stop | Out-Null
      $failure = ''
    } catch {
      $failure = $_.Exception.Message
    }
    if ($failure -notmatch 'Unauthorized|Forbidden') { break }
  }
  if ($failure) {
    throw "SharePoint refused the connection to $ConfigSiteUrl ($failure). Check that the account belongs to tenant $TenantId and has access to the site (a login persisted with -PersistLogin is reused without asking: see Get-PnPPersistedLogin and Disconnect-PnPOnline -ClearPersistedLogin), that the URL is right, and that the app registration $ClientId is consented for SharePoint in this tenant (otherwise pass -ClientId)."
  }
  $grantSite = {
    param([string]$Url, [string]$Level)
    $grants = @(Get-PnPAzureADAppSitePermission -Site $Url -Connection $connection -ErrorAction SilentlyContinue |
        Where-Object { $_.Apps.Id -contains $identity.clientId })
    $current = $grants | Sort-Object { ($_.Roles | ForEach-Object { $rank["$_".ToLower()] } | Measure-Object -Maximum).Maximum } -Descending | Select-Object -First 1
    $currentRank = if ($current) { ($current.Roles | ForEach-Object { $rank["$_".ToLower()] } | Measure-Object -Maximum).Maximum } else { 0 }
    if (-not $current) {
      Grant-PnPAzureADAppSitePermission -AppId $identity.clientId -DisplayName $FunctionAppName -Site $Url -Permissions $Level -Connection $connection -ErrorAction Stop | Out-Null
      Write-Done "${Level}: $Url (granted)"
    } elseif ($currentRank -lt $rank[$Level.ToLower()]) {
      Set-PnPAzureADAppSitePermission -Site $Url -PermissionId $current.Id -Permissions $Level -Connection $connection -ErrorAction Stop | Out-Null
      Write-Done "${Level}: $Url (raised from $($current.Roles -join ', '))"
    } else {
      Write-Info "$($current.Roles -join ', '): $Url (kept)"
    }
  }

  # Granting an app access to a site takes a site collection administrator of that site. On a
  # site that refuses the grant, the signed-in account is made one for the grant (through the
  # SharePoint admin site, which takes the SharePoint Administrator role) and removed again
  # afterwards. The preflight says so before anything is done.
  $adminConnection = $null
  $signedInUser = ''
  foreach ($url in $wanted.Keys) {
    $level = $wanted[$url]
    try {
      & $grantSite $url $level
      continue
    } catch {
      $refusal = "$($_.Exception.Message)"
      if ($refusal -notmatch 'access ?denied|unauthorized|forbidden|403') { throw }
    }
    if (-not $adminConnection) {
      $adminUrl = "https://$(([uri]$ConfigSiteUrl).Host -replace '\.sharepoint\.', '-admin.sharepoint.')"
      $adminConnection = Connect-PnPOnline -Url $adminUrl -Interactive -ClientId $ClientId -Tenant $TenantId -ReturnConnection
      $web = Get-PnPWeb -Includes CurrentUser -Connection $connection
      $signedInUser = ("$($web.CurrentUser.LoginName)" -split '\|')[-1]
      if (-not $signedInUser) { throw 'Could not read the signed-in account from SharePoint, so it cannot be made a site collection administrator. Add it in the SharePoint admin center.' }
    }
    try {
      Set-PnPTenantSite -Identity $url -Owners $signedInUser -Connection $adminConnection -ErrorAction Stop
    } catch {
      throw "Access denied when granting $level on ${url}, and $signedInUser could not be made a site collection administrator there ($($_.Exception.Message)). That takes an active SharePoint Administrator or Global Administrator role. Activate the role, or add the account in the SharePoint admin center (Active sites > the site > Membership > Site admins), and run again."
    }
    Write-Info "Made $signedInUser a site collection administrator of $url for the grant"
    try {
      $granted = $false
      foreach ($wait in 5, 15, 30, 60) {
        Start-Sleep -Seconds $wait
        try {
          & $grantSite $url $level
          $granted = $true
          break
        } catch {
          $refusal = "$($_.Exception.Message)"
        }
      }
      if (-not $granted) { throw "Granting $level on $url was still refused after the account was made a site collection administrator: $refusal" }
    } finally {
      try {
        $siteConnection = Connect-PnPOnline -Url $url -Interactive -ClientId $ClientId -Tenant $TenantId -ReturnConnection
        Remove-PnPSiteCollectionAdmin -Owners $signedInUser -Connection $siteConnection -ErrorAction Stop
        Write-Info "Removed $signedInUser as site collection administrator of $url"
      } catch {
        Write-Warning "Could not remove $signedInUser as site collection administrator of $url ($($_.Exception.Message))."
        $manualSteps.Add("Remove $signedInUser as site collection administrator of $url (SharePoint admin center > Active sites > the site > Membership), unless the account should stay.")
      }
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
    $catalogApp = Find-CatalogApp $TeamsAppId
  }
  if ($catalogApp) {
    # Never upload over an app that is not the notification app. The sender in the manifest
    # (webApplicationInfo.id) settles it when Graph returns it; otherwise the name must match.
    $newest = @($catalogApp.appDefinitions) | Sort-Object { [version]$_.version } -Descending | Select-Object -First 1
    $sender = "$($newest.authorization.clientAppId)"
    $name = if ($newest.displayName) { "$($newest.displayName)" } else { "$($catalogApp.displayName)" }
    if ($sender -and $sender -ne $identity.clientId) {
      throw "The Teams app $TeamsAppId in the app catalog is «$name», and its sender (webApplicationInfo.id) is $sender, not the notification identity $($identity.clientId). It is not the notification app, so nothing is uploaded. Pass -TeamsAppId with the app ID of the notification app."
    }
    if (-not $sender -and $name -ne $AppName) {
      throw "The Teams app $TeamsAppId in the app catalog is named «$name», not «$AppName», so it may not be the notification app and nothing is uploaded. Pass -TeamsAppId with the app ID of the notification app. To rename the app, run with -SkipTeamsAppUpload and upload the package in Teams admin center."
    }
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
    $created = Invoke-MgGraphRequest -Method POST -Uri 'v1.0/appCatalogs/teamsApps' -InputFilePath $zip.FullName -ContentType 'application/zip'
    # Graph has answered this call without an error while the app never showed up in the catalog,
    # so look the app up afterwards instead of trusting the answer.
    $listed = $null
    foreach ($wait in 0, 5, 10, 15) {
      Start-Sleep -Seconds $wait
      $listed = Find-CatalogApp $TeamsAppId
      if ($listed) { break }
    }
    if ($listed) {
      Write-Done "Published $AppName ($TeamsAppId) to the organization's app catalog (catalog ID $($listed.id))"
    } else {
      $answer = if ($created) { $created | ConvertTo-Json -Depth 5 -Compress } else { '(empty)' }
      Write-Warning "Graph accepted $AppName, but the app is not in the organization's app catalog. Answer from Graph: $answer"
      $manualSteps.Add("The upload with Graph did not put the app in the catalog. Upload $($zip.FullName) in Teams admin center (Teams apps > Manage apps > Upload new app).")
    }
    $manualSteps.Add("Teams admin center: create an app setup policy that installs «$AppName» for the recipients (without pinning it) and assign it to the same groups as in Varslingskonfigurasjon. Start with a test group.")
  }
}

# Remember the notification app's ID for later runs. Only this script reads the setting.
if ("$($appSettings['TEAMS_APP_MANIFEST_ID'])" -ne $TeamsAppId) {
  & az functionapp config appsettings set -n $FunctionAppName -g $ResourceGroupName --settings "TEAMS_APP_MANIFEST_ID=$TeamsAppId" -o none
  if ($LASTEXITCODE -ne 0) { throw "Could not store TEAMS_APP_MANIFEST_ID on $FunctionAppName" }
  Write-Info "Stored the notification app's ID in the app setting TEAMS_APP_MANIFEST_ID"
}
if ($opensOtherApp) {
  $manualSteps.Add("The notifications open the Teams app $deepLinkAppId (tab $deepLinkEntityId). Its tab must pass {subEntityId} to the page as the URL parameter driftsmelding, and the recipients need «$AppName» installed as well (README, step 4).")
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
