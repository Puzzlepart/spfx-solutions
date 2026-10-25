<#
.SYNOPSIS
  Builds the Teams app package (zip) for the notification service from the manifest template.

.DESCRIPTION
  Fills in the placeholders in ../teams-app/manifest.json and zips it together with the icons.
  The package is uploaded in Teams admin center (Teams apps > Manage apps > Upload new app) and
  rolled out with an app setup policy. Re-run with the same -AppId to produce a new version of the
  same app; bump -Version when re-uploading.

.PARAMETER NotificationClientId
  Client ID of the identity that sends the notifications (IdentityClientId from
  Deploy-NotificationInfrastructure.ps1). Becomes webApplicationInfo.id, which Teams requires to
  match the sender of activity notifications.

.PARAMETER AppId
  Teams app ID (a GUID). Generated when omitted; keep it for later updates of the same app.

.PARAMETER SiteUrl
  The intranet site the static tab opens, e.g. https://contoso.sharepoint.com

.PARAMETER ColorIcon / OutlineIcon
  Replace the default icons (192x192 color, 32x32 transparent outline) with the intranet's own.

.EXAMPLE
  .\New-NotificationTeamsApp.ps1 -NotificationClientId 360465b8-... -SiteUrl https://contoso.sharepoint.com -AppName "Intranett-varsler" -DeveloperName "Contoso"
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)] [string]$NotificationClientId,
  [Parameter(Mandatory = $true)] [string]$SiteUrl,
  [Parameter(Mandatory = $true)] [string]$AppName,
  [Parameter(Mandatory = $true)] [string]$DeveloperName,
  [string]$AppId = [guid]::NewGuid().ToString(),
  [string]$Version = '1.0.0',
  [string]$ColorIcon = '',
  [string]$OutlineIcon = '',
  [string]$TemplatePath = (Join-Path $PSScriptRoot '..\teams-app'),
  [string]$OutputPath = (Join-Path $PSScriptRoot '..\teams-app\dist')
)

$ErrorActionPreference = 'Stop'
$SiteUrl = $SiteUrl.TrimEnd('/')
$siteHost = ([uri]$SiteUrl).Host
if (-not $ColorIcon)   { $ColorIcon   = Join-Path $TemplatePath 'color.png' }
if (-not $OutlineIcon) { $OutlineIcon = Join-Path $TemplatePath 'outline.png' }

$manifest = Get-Content (Join-Path $TemplatePath 'manifest.json') -Raw -Encoding UTF8
$values = [ordered]@{
  '{{TEAMS_APP_ID}}'            = $AppId
  '{{NOTIFICATION_CLIENT_ID}}'  = $NotificationClientId
  '{{SITE_URL_ENCODED}}'        = [uri]::EscapeDataString($SiteUrl)
  '{{SITE_URL}}'                = $SiteUrl
  '{{SITE_HOST}}'               = $siteHost
  '{{APP_NAME}}'                = $AppName
  '{{DEVELOPER_NAME}}'          = $DeveloperName
}
foreach ($key in $values.Keys) { $manifest = $manifest.Replace($key, [string]$values[$key]) }
$manifest = $manifest -replace '"version":\s*"[^"]*"', "`"version`": `"$Version`""
$left = [regex]::Matches($manifest, '\{\{[A-Z_]+\}\}') | ForEach-Object { $_.Value } | Select-Object -Unique
if ($left) { throw "Unreplaced placeholders in manifest: $($left -join ', ')" }
$null = $manifest | ConvertFrom-Json   # validates JSON

$work = Join-Path ([System.IO.Path]::GetTempPath()) "teamsapp-$([guid]::NewGuid())"
New-Item -ItemType Directory -Path $work | Out-Null
try {
  [System.IO.File]::WriteAllText((Join-Path $work 'manifest.json'), $manifest, [System.Text.UTF8Encoding]::new($false))
  Copy-Item $ColorIcon (Join-Path $work 'color.png')
  Copy-Item $OutlineIcon (Join-Path $work 'outline.png')

  New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
  $safeName = ($AppName -replace '[^\w\-]+', '-').Trim('-').ToLower()
  $zip = Join-Path (Resolve-Path $OutputPath) "$safeName-$Version.zip"
  if (Test-Path $zip) { Remove-Item $zip }
  Compress-Archive -Path (Join-Path $work '*') -DestinationPath $zip
} finally {
  Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host "Teams app package: $zip" -ForegroundColor Green
[pscustomobject]@{
  Package              = $zip
  TeamsAppId           = $AppId
  Version              = $Version
  WebApplicationInfoId = $NotificationClientId
  TabUrl               = $SiteUrl
} | Format-List
Write-Host 'Upload in Teams admin center: Teams apps > Manage apps > Upload new app. Keep TeamsAppId for later updates.'
Write-Host 'Point the notifications at this app (its tab opens the intranet):'
Write-Host "  az functionapp config appsettings set -g <resource-group> -n <function-app> --settings TEAMS_DEEPLINK_APP_ID=$AppId TEAMS_DEEPLINK_ENTITY_ID=intranet\"
