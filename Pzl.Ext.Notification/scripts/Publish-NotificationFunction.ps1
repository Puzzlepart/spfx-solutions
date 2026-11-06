<#
.SYNOPSIS
  Builds the notification function and deploys it to the function app with Azure CLI zip deploy.

.DESCRIPTION
  Builds locally and ships a complete package, so nothing is built in Azure:
    1. npm ci + npm run build in the function folder (TypeScript -> dist/)
    2. Stages host.json, package.json, package-lock.json and dist/ in a temp folder
    3. npm ci --omit=dev in the staging folder (production node_modules only)
    4. Zips the staging folder and deploys it with az functionapp deployment source config-zip
       --build-remote false
  Remote (Oryx) builds are avoided on purpose: Oryx runs "npm run build" when package.json has a
  build script, which fails because src/ is not part of the package.

  Requires Azure CLI logged in to the subscription that holds the function app, and Node 20+.
  Azure Functions Core Tools is not required.

.PARAMETER FunctionPath
  Path to the function project. Defaults to ../function next to this script.

.PARAMETER PackageOnly
  Build the zip and print its path without deploying (for inspection or hand-over).

.EXAMPLE
  .\Publish-NotificationFunction.ps1 -ResourceGroupName rg-intranet-notification -FunctionAppName func-intranet-notification
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)] [string]$ResourceGroupName,
  [Parameter(Mandatory = $true)] [string]$FunctionAppName,
  [string]$FunctionPath = (Join-Path $PSScriptRoot '..\function'),
  [switch]$SkipBuild,
  [switch]$PackageOnly,
  [int]$TimeoutSeconds = 900
)

$ErrorActionPreference = 'Stop'
$FunctionPath = (Resolve-Path $FunctionPath).Path

if (-not $SkipBuild) {
  Write-Host "== Build ($FunctionPath)" -ForegroundColor Cyan
  Push-Location $FunctionPath
  try {
    & npm ci --no-audit --no-fund; if ($LASTEXITCODE -ne 0) { throw 'npm ci failed' }
    & npm run build;               if ($LASTEXITCODE -ne 0) { throw 'npm run build failed' }
  } finally { Pop-Location }
}
if (-not (Test-Path (Join-Path $FunctionPath 'dist\src\index.js'))) { throw "Build output missing: $FunctionPath\dist\src\index.js" }

Write-Host '== Package' -ForegroundColor Cyan
$stamp = Get-Date -Format yyyyMMdd-HHmmss
$staging = Join-Path ([System.IO.Path]::GetTempPath()) "$FunctionAppName-$stamp"
$zip = "$staging.zip"
New-Item -ItemType Directory -Path $staging | Out-Null
try {
  foreach ($item in 'host.json', 'package.json', 'package-lock.json') { Copy-Item (Join-Path $FunctionPath $item) $staging }
  Copy-Item (Join-Path $FunctionPath 'dist') (Join-Path $staging 'dist') -Recurse
  Push-Location $staging
  try {
    & npm ci --omit=dev --no-audit --no-fund --ignore-scripts; if ($LASTEXITCODE -ne 0) { throw 'npm ci --omit=dev failed in staging folder' }
  } finally { Pop-Location }
  Remove-Item (Join-Path $staging 'package-lock.json')
  Compress-Archive -Path (Join-Path $staging '*') -DestinationPath $zip -CompressionLevel Optimal
} finally {
  Remove-Item $staging -Recurse -Force -ErrorAction SilentlyContinue
}
$sizeMb = [math]::Round((Get-Item $zip).Length / 1MB, 1)
Write-Host "   $zip ($sizeMb MB: host.json, package.json, dist/, production node_modules/)"

if ($PackageOnly) {
  Write-Host 'PackageOnly: not deploying.' -ForegroundColor Yellow
  return $zip
}

Write-Host "== Deploy to $FunctionAppName (no remote build)" -ForegroundColor Cyan
& az functionapp deployment source config-zip -g $ResourceGroupName -n $FunctionAppName --src $zip --build-remote false --timeout $TimeoutSeconds -o none
if ($LASTEXITCODE -ne 0) { throw "Deployment failed. Package kept at $zip" }
Remove-Item $zip -ErrorAction SilentlyContinue

Write-Host '== Functions in the app' -ForegroundColor Cyan
& az functionapp function list -g $ResourceGroupName -n $FunctionAppName --query '[].{name:name, trigger:config.bindings[0].type}' -o table

Write-Host ''
Write-Host 'Deployed. Watch the first runs with (logs arrive after a minute or two):' -ForegroundColor Green
Write-Host "  .\Get-NotificationLog.ps1 -ResourceGroupName $ResourceGroupName -FunctionAppName $FunctionAppName"
Write-Host 'Set DRY_RUN=false when the dry run looks right:'
Write-Host "  az functionapp config appsettings set -g $ResourceGroupName -n $FunctionAppName --settings DRY_RUN=false"
