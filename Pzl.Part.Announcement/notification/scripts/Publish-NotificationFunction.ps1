<#
.SYNOPSIS
  Builds the notification function and deploys it to the function app with Azure CLI zip deploy.

.DESCRIPTION
  Runs npm ci and npm run build in the function folder, packages host.json, package.json,
  package-lock.json and dist/ into a zip, and deploys it with a remote build (npm install runs
  in Azure, so node_modules is not uploaded). Azure Functions Core Tools is not required.

  Requires Azure CLI logged in to the subscription that holds the function app, and Node 20+.

.PARAMETER FunctionPath
  Path to the function project. Defaults to ../function next to this script.

.EXAMPLE
  .\Publish-NotificationFunction.ps1 -ResourceGroupName rg-intranet-notification -FunctionAppName func-intranet-notification
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)] [string]$ResourceGroupName,
  [Parameter(Mandatory = $true)] [string]$FunctionAppName,
  [string]$FunctionPath = (Join-Path $PSScriptRoot '..\function'),
  [switch]$SkipBuild,
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
$zip = Join-Path ([System.IO.Path]::GetTempPath()) "$FunctionAppName-$(Get-Date -Format yyyyMMdd-HHmmss).zip"
Push-Location $FunctionPath
try {
  Compress-Archive -Path 'host.json', 'package.json', 'package-lock.json', 'dist' -DestinationPath $zip -CompressionLevel Optimal
} finally { Pop-Location }
Write-Host "   $zip ($([math]::Round((Get-Item $zip).Length / 1KB)) KB)"

Write-Host "== Deploy to $FunctionAppName (remote build)" -ForegroundColor Cyan
& az functionapp deployment source config-zip -g $ResourceGroupName -n $FunctionAppName --src $zip --build-remote true --timeout $TimeoutSeconds -o none
if ($LASTEXITCODE -ne 0) { throw 'Deployment failed' }
Remove-Item $zip -ErrorAction SilentlyContinue

Write-Host '== Functions in the app' -ForegroundColor Cyan
& az functionapp function list -g $ResourceGroupName -n $FunctionAppName --query '[].{name:name, trigger:config.bindings[0].type}' -o table

Write-Host ''
Write-Host 'Deployed. Watch the first runs with:' -ForegroundColor Green
Write-Host "  az webapp log tail -g $ResourceGroupName -n $FunctionAppName"
Write-Host "Set DRY_RUN=false when the dry run looks right:"
Write-Host "  az functionapp config appsettings set -g $ResourceGroupName -n $FunctionAppName --settings DRY_RUN=false"
