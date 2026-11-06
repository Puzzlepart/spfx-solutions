<#
.SYNOPSIS
  Shows recent log lines from the notification function (Application Insights).

.DESCRIPTION
  Flex Consumption function apps have no Kudu log stream, so "az webapp log tail" fails with 404.
  This script queries the traces table in Application Insights instead. Logs arrive with a delay
  of one to a few minutes.

.PARAMETER AppInsightsName
  Defaults to "<FunctionAppName>-insights", the name Deploy-NotificationInfrastructure.ps1 uses.

.EXAMPLE
  .\Get-NotificationLog.ps1 -ResourceGroupName rg-intranet-notification -FunctionAppName func-intranet-notification -Minutes 30
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)] [string]$ResourceGroupName,
  [Parameter(Mandatory = $true)] [string]$FunctionAppName,
  [string]$AppInsightsName = '',
  [int]$Minutes = 30,
  [int]$Take = 100,
  [switch]$IncludeHostMessages
)

$ErrorActionPreference = 'Stop'
if (-not $AppInsightsName) { $AppInsightsName = "$FunctionAppName-insights" }
$filter = if ($IncludeHostMessages) { '' } else { "| where message !startswith 'Host Status' and message !startswith 'Trigger Details' and message !startswith 'Executing '" }
$query = "traces | where timestamp > ago(${Minutes}m) $filter | project timestamp, severityLevel, message | order by timestamp desc | take $Take"

$json = & az monitor app-insights query --app $AppInsightsName -g $ResourceGroupName --analytics-query $query -o json 2>$null
if ($LASTEXITCODE -ne 0 -or -not $json) { throw "Query failed. Check that Application Insights '$AppInsightsName' exists in '$ResourceGroupName' and that you are logged in (az login)." }
$rows = ($json | ConvertFrom-Json).tables[0].rows
if (-not $rows) { Write-Host "No log lines in the last $Minutes minutes. The timer runs every minute; logs arrive with a delay."; return }

$levels = @{ 0 = 'Verbose'; 1 = 'Info'; 2 = 'Warn'; 3 = 'Error'; 4 = 'Critical' }
[array]::Reverse($rows)
foreach ($r in $rows) {
  $ts = ([datetime]$r[0]).ToLocalTime().ToString('HH:mm:ss')
  $lvl = $levels[[int]$r[1]]
  $color = switch ($lvl) { 'Error' { 'Red' } 'Critical' { 'Red' } 'Warn' { 'Yellow' } default { 'Gray' } }
  Write-Host ("{0} {1,-8} {2}" -f $ts, $lvl, $r[2]) -ForegroundColor $color
}
