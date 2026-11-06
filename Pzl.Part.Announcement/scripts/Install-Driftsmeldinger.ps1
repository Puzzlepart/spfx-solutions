<#
.SYNOPSIS
  Installs or upgrades Driftsmeldinger (Pzl.Part.Announcement) in a tenant, and optionally
  the SharePoint part of the Teams notification add-on (Pzl.Ext.Notification).

.DESCRIPTION
  Idempotent: run it for a new installation and for every upgrade. Each step checks what is
  there and only changes what is missing or outdated.

    1. Uploads and publishes pzl-announcement.sppkg to the tenant app catalog (replaces an
       existing version; same solution ID as 1.x, so 1.x is upgraded in place).
    2. Updates the app on the list site if it was added there from a site collection.
    3. Applies Templates/Announcement.xml to the list site (creates the list, or adds Status,
       the new level choices and the column descriptions to an existing one).
    4. Upgrade from 1.x: removes the AnnouncementDeepLink custom action (component no longer
       in the package) from the list site and the notification site.
    5. Optional (-BannerSiteUrls, -BannerHubSiteUrls or -BannerEverywhere): creates or updates
       the banner's row in the Tenant Wide Extensions list. Existing properties that are not
       given (such as popoverFields) are kept.
    6. Optional (-EnableTeamsNotifications): uploads pzl-ext-notification.sppkg and applies
       AnnouncementFields.xml (list site), NotificationConfig.xml and DeepLinkCustomAction.xml
       (notification site). The Azure, Entra and Teams part of the add-on is not covered, see
       Pzl.Ext.Notification/README.md.

  Requires PnP.PowerShell (not installed by this script) and SharePoint Administrator rights
  (tenant app catalog). Build the packages first with `npm run package` in each project.
  Supports -WhatIf.

.PARAMETER SiteUrl
  Site that hosts (or will host) the central Driftsmeldinger list, normally the root site.

.PARAMETER ClientId
  Client ID of the Entra app registration used by Connect-PnPOnline -Interactive. Defaults to
  the Prosjektportalen (PP365) PnP app, da6c31a6-b557-4ac3-9994-7315da06ea3a. Pass another
  if the tenant uses its own app registration.

.PARAMETER AppCatalogUrl
  Tenant app catalog site. Looked up with Get-PnPTenantAppCatalogUrl when empty.

.PARAMETER PackagePath
  Path to pzl-announcement.sppkg. Defaults to the build output of this project.

.PARAMETER SkipPackage
  Do not upload any package (only templates, cleanup and banner row).

.PARAMETER BannerSiteUrls
  Absolute site URLs where the banner is shown (targeting property siteUrls).

.PARAMETER BannerHubSiteUrls
  Hub site URLs or hub site IDs where the banner is shown (targeting property hubSiteUrls).

.PARAMETER BannerEverywhere
  Show the banner on every site (targeting property everywhere).

.PARAMETER BannerTargetAudience
  Filter the banner on the list's audience targeting column.

.PARAMETER EnableTeamsNotifications
  Also install the SharePoint part of the Teams notification add-on.

.PARAMETER NotificationSiteUrl
  Site that hosts the Varslingskonfigurasjon list and that the notification Teams tab opens
  (gets the NotificationDeepLink custom action). Defaults to SiteUrl.

.PARAMETER NotificationPackagePath
  Path to pzl-ext-notification.sppkg. Defaults to the build output of Pzl.Ext.Notification.

.EXAMPLE
  .\Install-Driftsmeldinger.ps1 -SiteUrl https://contoso.sharepoint.com

.EXAMPLE
  .\Install-Driftsmeldinger.ps1 -SiteUrl https://contoso.sharepoint.com `
    -BannerHubSiteUrls https://contoso.sharepoint.com/sites/hub -EnableTeamsNotifications -WhatIf
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
  [Parameter(Mandatory = $true)] [string]$SiteUrl,
  [string]$ClientId = 'da6c31a6-b557-4ac3-9994-7315da06ea3a',
  [string]$AppCatalogUrl = '',
  [string]$PackagePath = (Join-Path $PSScriptRoot '..\sharepoint\solution\pzl-announcement.sppkg'),
  [switch]$SkipPackage,
  [string[]]$BannerSiteUrls = @(),
  [string[]]$BannerHubSiteUrls = @(),
  [switch]$BannerEverywhere,
  [switch]$BannerTargetAudience,
  [switch]$EnableTeamsNotifications,
  [string]$NotificationSiteUrl = '',
  [string]$NotificationPackagePath = (Join-Path $PSScriptRoot '..\..\Pzl.Ext.Notification\sharepoint\solution\pzl-ext-notification.sppkg')
)

$ErrorActionPreference = 'Stop'

# Component IDs (unchanged since 1.x, except the deep link customizer)
$BannerComponentId = 'c0d1edf4-db39-4feb-8052-a4c27ffe1954'
$OldDeepLinkComponentId = '9f3d6a2c-7b14-4e8f-a5c0-3d2e1f6b8a47'
$TenantWideExtensionsList = 'Tenant Wide Extensions'

$TemplateFolder = (Resolve-Path (Join-Path $PSScriptRoot '..\Templates')).Path
$NotificationTemplateFolder = Join-Path $PSScriptRoot '..\..\Pzl.Ext.Notification\Templates'

$SiteUrl = $SiteUrl.TrimEnd('/')
if (-not $NotificationSiteUrl) { $NotificationSiteUrl = $SiteUrl }
$NotificationSiteUrl = $NotificationSiteUrl.TrimEnd('/')

$manualSteps = [System.Collections.Generic.List[string]]::new()

function Write-Step([string]$Text) { Write-Host "`n== $Text" -ForegroundColor Cyan }
function Write-Done([string]$Text) { Write-Host "   $Text" -ForegroundColor Green }
function Write-Info([string]$Text) { Write-Host "   $Text" }

#region PnP.PowerShell

Write-Step 'PnP.PowerShell'
$module = Get-Module -ListAvailable -Name PnP.PowerShell | Sort-Object Version -Descending | Select-Object -First 1
if (-not $module) {
  throw 'PnP.PowerShell is not installed. Install it yourself (this script does not), then run the script again.'
}
Import-Module $module.Path -ErrorAction Stop
Write-Done "PnP.PowerShell $($module.Version)"

# PnP.PowerShell 3.4.x does not resolve CSOM assemblies that are not loaded at import, and the
# templates use localization, which makes PnP.Framework load Microsoft.SharePoint.Client.UserProfiles.
# Pre-loading them is harmless on other versions. See README, "Known issue with PnP.PowerShell 3.4.x".
if ($module.Version -ge [version]'3.4.0') {
  $core = Join-Path $module.ModuleBase 'Core'
  if (Test-Path $core) {
    Get-ChildItem $core -Filter 'Microsoft.*SharePoint*.dll' |
      ForEach-Object { [System.Reflection.Assembly]::LoadFrom($_.FullName) | Out-Null }
    Write-Info 'Pre-loaded the CSOM assemblies (workaround for PnP.PowerShell 3.4.x)'
  }
}

#endregion

#region Connections

$connections = @{}
function Get-Connection([string]$Url) {
  $key = $Url.TrimEnd('/').ToLowerInvariant()
  if (-not $connections.ContainsKey($key)) {
    Write-Info "Connecting to $Url"
    $connections[$key] = Connect-PnPOnline -Url $Url -Interactive -ClientId $ClientId -ReturnConnection
  }
  return $connections[$key]
}

function Invoke-Template([string]$Path, [string]$Url) {
  $name = Split-Path $Path -Leaf
  if (-not (Test-Path $Path)) { throw "Template not found: $Path" }
  if ($PSCmdlet.ShouldProcess($Url, "Apply $name")) {
    Invoke-PnPSiteTemplate -Path (Resolve-Path $Path).Path -Connection (Get-Connection $Url)
    Write-Done "Applied $name to $Url"
  }
}

$siteConnection = Get-Connection $SiteUrl
if (-not $AppCatalogUrl) {
  $AppCatalogUrl = Get-PnPTenantAppCatalogUrl -Connection $siteConnection
  if (-not $AppCatalogUrl) { throw 'The tenant has no app catalog. Create it, or pass -AppCatalogUrl.' }
}
$AppCatalogUrl = $AppCatalogUrl.TrimEnd('/')
Write-Info "Tenant app catalog: $AppCatalogUrl"

#endregion

#region Current state

Write-Step 'Current state'
$list = Get-PnPList -Identity 'Lists/Driftsmeldinger' -Connection $siteConnection -ErrorAction SilentlyContinue
if (-not $list) {
  $mode = 'new installation'
  Write-Info "No Driftsmeldinger list on $SiteUrl"
} else {
  $statusField = Get-PnPField -List $list -Identity 'PzlStatus' -Connection $siteConnection -ErrorAction SilentlyContinue
  $mode = if ($statusField) { 'update of 2.x' } else { 'upgrade from 1.x' }
  Write-Info "Driftsmeldinger list found ($($list.ItemCount) items)"
}
Write-Done "Mode: $mode"

#endregion

#region Packages

function Publish-Package([string]$Path, [string]$Title) {
  if (-not (Test-Path $Path)) {
    throw "Package not found: $Path. Build it with 'npm run package' in the project folder, or pass the path."
  }
  $resolved = (Resolve-Path $Path).Path
  if ($PSCmdlet.ShouldProcess($AppCatalogUrl, "Upload and publish $(Split-Path $resolved -Leaf)")) {
    $app = Add-PnPApp -Path $resolved -Scope Tenant -Overwrite -Publish -SkipFeatureDeployment -Connection $siteConnection
    Write-Done "$Title published to the tenant app catalog (app ID $($app.Id))"
    return $app
  }
}

function Update-SiteApp($App, [string]$Url, [string]$Title) {
  if (-not $App) { return }
  $connection = Get-Connection $Url
  $installed = Get-PnPApp -Identity $App.Id -Scope Tenant -Connection $connection -ErrorAction SilentlyContinue
  if ($installed -and $installed.InstalledVersion -and $installed.CanUpgrade) {
    if ($PSCmdlet.ShouldProcess($Url, "Update $Title from $($installed.InstalledVersion)")) {
      Update-PnPApp -Identity $App.Id -Scope Tenant -Connection $connection
      Write-Done "Updated $Title on $Url"
    }
  }
}

if ($SkipPackage) {
  Write-Step 'Packages (skipped)'
} else {
  Write-Step 'Driftsmeldinger package'
  $announcementApp = Publish-Package $PackagePath 'Driftsmeldinger'
  Update-SiteApp $announcementApp $SiteUrl 'Driftsmeldinger'
}

#endregion

#region List

Write-Step 'Driftsmeldinger list'
Invoke-Template (Join-Path $TemplateFolder 'Announcement.xml') $SiteUrl
if ($mode -eq 'upgrade from 1.x') {
  $manualSteps.Add('Existing items have no status (shown as Pågår) and may have the 1.x levels Varsel/Normal (shown as Kritisk/Informasjon). Editors pick a new level and a status the next time they edit them.')
  $manualSteps.Add('The columns «Varsle i Teams» and «Varslet i Teams» stay on the list. Without the notification add-on they do nothing; delete them in the list settings if they are not wanted.')
}

#endregion

#region Cleanup from 1.x

Write-Step 'Cleanup from 1.x'
$cleanupSites = @($SiteUrl, $NotificationSiteUrl) | Select-Object -Unique
$removed = 0
foreach ($url in $cleanupSites) {
  $connection = Get-Connection $url
  foreach ($scope in 'Site', 'Web') {
    $actions = Get-PnPCustomAction -Scope $scope -Connection $connection |
      Where-Object { "$($_.ClientSideComponentId)" -eq $OldDeepLinkComponentId }
    foreach ($action in $actions) {
      if ($PSCmdlet.ShouldProcess("$url ($scope)", "Remove custom action $($action.Title)")) {
        Remove-PnPCustomAction -Identity $action.Id -Scope $scope -Force -Connection $connection
        Write-Done "Removed the AnnouncementDeepLink custom action from $url ($scope)"
        $removed++
      }
    }
  }
}
if (-not $removed) { Write-Info 'Nothing to remove' }

#endregion

#region Banner

$bannerTargeting = $BannerSiteUrls.Count -or $BannerHubSiteUrls.Count -or $BannerEverywhere
if ($bannerTargeting) {
  Write-Step 'Banner (Tenant Wide Extensions)'
  $catalogConnection = Get-Connection $AppCatalogUrl
  $rows = Get-PnPListItem -List $TenantWideExtensionsList -Connection $catalogConnection -PageSize 500 |
    Where-Object { "$($_['TenantWideExtensionComponentId'])" -eq $BannerComponentId }
  $row = $rows | Select-Object -First 1
  if (@($rows).Count -gt 1) { Write-Warning "Found $(@($rows).Count) banner rows; updating the first (ID $($row.Id)) only." }

  $properties = [ordered]@{}
  if ($row -and $row['TenantWideExtensionComponentProperties']) {
    $existing = $row['TenantWideExtensionComponentProperties'] | ConvertFrom-Json
    $existing.PSObject.Properties | ForEach-Object { $properties[$_.Name] = $_.Value }
  }
  $properties['listSiteUrl'] = $SiteUrl
  $properties['targetAudience'] = [bool]$BannerTargetAudience
  if ($BannerSiteUrls.Count) { $properties['siteUrls'] = @($BannerSiteUrls | ForEach-Object { $_.TrimEnd('/') }) }
  if ($BannerHubSiteUrls.Count) { $properties['hubSiteUrls'] = @($BannerHubSiteUrls | ForEach-Object { $_.TrimEnd('/') }) }
  if ($BannerEverywhere) { $properties['everywhere'] = $true }
  $json = $properties | ConvertTo-Json -Compress -Depth 5

  $values = @{
    Title                                  = 'AnnouncementBanner'
    TenantWideExtensionComponentId         = $BannerComponentId
    TenantWideExtensionLocation            = 'ClientSideExtension.ApplicationCustomizer'
    TenantWideExtensionComponentProperties = $json
    TenantWideExtensionSequence            = 110
    TenantWideExtensionListTemplate        = 0
    TenantWideExtensionDisabled            = $false
  }
  if ($row) {
    if ($PSCmdlet.ShouldProcess("$TenantWideExtensionsList, item $($row.Id)", "Update banner properties to $json")) {
      Set-PnPListItem -List $TenantWideExtensionsList -Identity $row.Id -Values $values -Connection $catalogConnection | Out-Null
      Write-Done "Updated the banner row (ID $($row.Id)): $json"
    }
  } elseif ($PSCmdlet.ShouldProcess($TenantWideExtensionsList, "Add banner row with $json")) {
    $item = Add-PnPListItem -List $TenantWideExtensionsList -Values $values -Connection $catalogConnection
    Write-Done "Added the banner row (ID $($item.Id)): $json"
  }
} else {
  Write-Step 'Banner (skipped, no targeting given)'
}

#endregion

#region Teams notification add-on (SharePoint part)

if ($EnableTeamsNotifications) {
  Write-Step 'Teams notification add-on (SharePoint part)'
  if (-not (Test-Path $NotificationTemplateFolder)) { throw "Pzl.Ext.Notification not found next to this project: $NotificationTemplateFolder" }
  if (-not $SkipPackage) {
    $notificationApp = Publish-Package $NotificationPackagePath 'Driftsmeldinger - Teams-varsling'
    Update-SiteApp $notificationApp $NotificationSiteUrl 'Driftsmeldinger - Teams-varsling'
  }
  # After Announcement.xml: it repeats the content type with the Teams fields added
  Invoke-Template (Join-Path $NotificationTemplateFolder 'AnnouncementFields.xml') $SiteUrl
  Invoke-Template (Join-Path $NotificationTemplateFolder 'NotificationConfig.xml') $NotificationSiteUrl
  Invoke-Template (Join-Path $NotificationTemplateFolder 'DeepLinkCustomAction.xml') $NotificationSiteUrl
  $manualSteps.Add("Add one row per site to Varslingskonfigurasjon on $NotificationSiteUrl.")
  $manualSteps.Add('Turn on «Varsle i Teams» under «Felter i skjema for ny driftsmelding» in the web part properties if editors register announcements there.')
  $manualSteps.Add('Azure, Entra and Teams part of the notification add-on: Pzl.Ext.Notification/README.md, steps 1, 2, 4 and 5.')
}

#endregion

#region Summary

Write-Step 'Done'
Write-Done "Driftsmeldinger ($mode) on $SiteUrl"
$manualSteps.Insert(0, "Add the Driftsmeldinger web part to a page. On other sites than $SiteUrl, set «Område for driftsmeldinger» to $SiteUrl.")
if (-not $bannerTargeting) {
  $manualSteps.Add('Banner: run again with -BannerSiteUrls, -BannerHubSiteUrls or -BannerEverywhere, or use Templates/BannerCustomAction.xml for one site.')
}
Write-Host "`nManual steps:" -ForegroundColor Yellow
$manualSteps | ForEach-Object { Write-Host "  - $_" }

#endregion
