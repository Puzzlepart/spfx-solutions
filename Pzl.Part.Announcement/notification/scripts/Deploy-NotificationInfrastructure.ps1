<#
.SYNOPSIS
  Creates the Azure resources for the Teams notification service and configures the function app.

.DESCRIPTION
  Idempotent: every resource is created only if it does not exist, so the script can be re-run.
  Creates, in order:
    1. Resource group
    2. User-assigned managed identity (the identity the service runs as)
    3. Storage account (Functions runtime, deployment package and the NotificationState table)
    4. Application Insights
    5. Role assignments for the identity on the storage account
    6. Function app (Flex Consumption, Node 22) using the identity for storage and deployment
    7. App settings for the service (identity-based storage, no connection strings)

  Requires Azure CLI 2.70 or later, logged in with rights to create resources and role assignments
  in the subscription (Owner, or Contributor + User Access Administrator).

  After this script: run Grant-NotificationPermissions.ps1 with the printed principal ID,
  grant site permissions with Grant-PnPAzureADAppSitePermission, create the Teams app,
  then deploy the code with Publish-NotificationFunction.ps1. See ../README.md.

.PARAMETER Location
  Region for all resources. Must support Flex Consumption: az functionapp list-flexconsumption-locations

.PARAMETER TeamsDeepLinkAppId
  Teams app ID the notifications open (Graph requires links to be Teams deep links). Usually the
  notification app itself, whose static tab "intranet" opens the site. It is created in step 4, so this
  is normally set afterwards: az functionapp config appsettings set ... --settings TEAMS_DEEPLINK_APP_ID=<id>

.PARAMETER DryRun
  Value for the DRY_RUN app setting. Keep $true until the first test run looks right.

.EXAMPLE
  .\Deploy-NotificationInfrastructure.ps1 -SubscriptionId <guid> -ResourceGroupName rg-intranet-notification `
    -FunctionAppName func-intranet-notification -StorageAccountName stintranetnotify01 `
    -ConfigSiteUrl https://contoso.sharepoint.com
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)] [string]$SubscriptionId,
  [Parameter(Mandatory = $true)] [string]$ResourceGroupName,
  [string]$Location = 'westeurope',
  [Parameter(Mandatory = $true)] [string]$FunctionAppName,
  [Parameter(Mandatory = $true)] [ValidatePattern('^[a-z0-9]{3,24}$')] [string]$StorageAccountName,
  [string]$IdentityName = '',
  [string]$AppInsightsName = '',
  [Parameter(Mandatory = $true)] [string]$ConfigSiteUrl,
  [string]$ConfigListTitle = 'Varslingskonfigurasjon',
  [string]$AnnouncementListTitle = 'Driftsmeldinger',
  [string]$AlwaysNotifySeverities = 'Varsel,Alert',
  [string]$NewsLookbackDays = '7',
  [string]$TeamsDeepLinkAppId = '',
  [string]$TeamsDeepLinkEntityId = 'intranet',
  [bool]$DryRun = $true
)

$ErrorActionPreference = 'Stop'
if (-not $IdentityName) { $IdentityName = "$FunctionAppName-identity" }
if (-not $AppInsightsName) { $AppInsightsName = "$FunctionAppName-insights" }

function Invoke-Az {
  # Runs az, throws on failure, returns parsed JSON (or $null for empty output).
  # az writes warnings to stderr; they are kept apart from stdout so they never end up in the JSON.
  param([Parameter(ValueFromRemainingArguments = $true)] [string[]]$Arguments)
  $raw = & az @Arguments -o json 2>&1
  $exit = $LASTEXITCODE
  $stderr = @($raw | Where-Object { $_ -is [System.Management.Automation.ErrorRecord] } | ForEach-Object { "$_" })
  $stdout = @($raw | Where-Object { $_ -isnot [System.Management.Automation.ErrorRecord] } | ForEach-Object { "$_" })
  foreach ($line in $stderr) { if ($line -match '^WARNING') { Write-Verbose $line } }
  if ($exit -ne 0) { throw "az $($Arguments[0..1] -join ' ') failed:`n$(($stderr + $stdout) -join "`n")" }
  $text = ($stdout -join "`n").Trim()
  # Keep only the JSON part in case a notice was printed on stdout before it
  $idx = @($text.IndexOf('{'), $text.IndexOf('[')) | Where-Object { $_ -ge 0 } | Sort-Object | Select-Object -First 1
  if ($null -eq $idx) { return $null }
  return ($text.Substring($idx) | ConvertFrom-Json)
}

function Test-Az {
  # Returns $true if the az command succeeds (used for "exists?" checks), without printing errors
  param([Parameter(ValueFromRemainingArguments = $true)] [string[]]$Arguments)
  $null = & az @Arguments -o none 2>$null
  return ($LASTEXITCODE -eq 0)
}

Write-Host "== Subscription" -ForegroundColor Cyan
Invoke-Az account set --subscription $SubscriptionId | Out-Null
$account = Invoke-Az account show
Write-Host "   $($account.name) ($($account.id)), tenant $($account.tenantId), user $($account.user.name)"
& az config set extension.use_dynamic_install=yes_without_prompt -o none 2>$null

Write-Host "== Resource group $ResourceGroupName ($Location)" -ForegroundColor Cyan
if (-not (Test-Az group show -n $ResourceGroupName)) {
  Invoke-Az group create -n $ResourceGroupName -l $Location | Out-Null; Write-Host '   created'
} else { Write-Host '   exists' }

Write-Host "== Managed identity $IdentityName" -ForegroundColor Cyan
if (-not (Test-Az identity show -n $IdentityName -g $ResourceGroupName)) {
  $identity = Invoke-Az identity create -n $IdentityName -g $ResourceGroupName -l $Location; Write-Host '   created'
} else { $identity = Invoke-Az identity show -n $IdentityName -g $ResourceGroupName; Write-Host '   exists' }
Write-Host "   clientId    $($identity.clientId)"
Write-Host "   principalId $($identity.principalId)"

Write-Host "== Storage account $StorageAccountName" -ForegroundColor Cyan
if (-not (Test-Az storage account show -n $StorageAccountName -g $ResourceGroupName)) {
  $storage = Invoke-Az storage account create -n $StorageAccountName -g $ResourceGroupName -l $Location `
    --sku Standard_LRS --kind StorageV2 --min-tls-version TLS1_2 --allow-blob-public-access false
  Write-Host '   created'
} else { $storage = Invoke-Az storage account show -n $StorageAccountName -g $ResourceGroupName; Write-Host '   exists' }

Write-Host "== Application Insights $AppInsightsName" -ForegroundColor Cyan
if (-not (Test-Az monitor app-insights component show --app $AppInsightsName -g $ResourceGroupName)) {
  Invoke-Az monitor app-insights component create --app $AppInsightsName -g $ResourceGroupName -l $Location --kind web --application-type web | Out-Null
  Write-Host '   created'
} else { Write-Host '   exists' }

Write-Host "== Role assignments for the identity on the storage account" -ForegroundColor Cyan
$newAssignments = 0
foreach ($role in 'Storage Blob Data Owner', 'Storage Queue Data Contributor', 'Storage Table Data Contributor') {
  $existing = Invoke-Az role assignment list --assignee-object-id $identity.principalId --role $role --scope $storage.id
  if ($existing.Count -eq 0) {
    Invoke-Az role assignment create --assignee-object-id $identity.principalId --assignee-principal-type ServicePrincipal --role $role --scope $storage.id | Out-Null
    Write-Host "   assigned: $role"; $newAssignments++
  } else { Write-Host "   exists:   $role" }
}
if ($newAssignments -gt 0) { Write-Host '   waiting 30 s for role assignments to propagate'; Start-Sleep -Seconds 30 }

Write-Host "== Function app $FunctionAppName (Flex Consumption, Node 22)" -ForegroundColor Cyan
if (-not (Test-Az functionapp show -n $FunctionAppName -g $ResourceGroupName)) {
  Invoke-Az functionapp create -n $FunctionAppName -g $ResourceGroupName --storage-account $StorageAccountName `
    --flexconsumption-location $Location --runtime node --runtime-version 22 `
    --assign-identity $identity.id --app-insights $AppInsightsName `
    --deployment-storage-auth-type UserAssignedIdentity --deployment-storage-auth-value $identity.id | Out-Null
  Write-Host '   created'
} else { Write-Host '   exists' }

Write-Host "== App settings" -ForegroundColor Cyan
$settings = @(
  "AzureWebJobsStorage__accountName=$StorageAccountName",
  "AzureWebJobsStorage__clientId=$($identity.clientId)",
  "AzureWebJobsStorage__credential=managedidentity",
  "AZURE_CLIENT_ID=$($identity.clientId)",
  "CONFIG_SITE_URL=$($ConfigSiteUrl.TrimEnd('/'))",
  "CONFIG_LIST_TITLE=$ConfigListTitle",
  "ANNOUNCEMENT_LIST_TITLE=$AnnouncementListTitle",
  "ALWAYS_NOTIFY_SEVERITIES=$AlwaysNotifySeverities",
  "NEWS_LOOKBACK_DAYS=$NewsLookbackDays",
  "TEAMS_DEEPLINK_APP_ID=$TeamsDeepLinkAppId",
  "TEAMS_DEEPLINK_ENTITY_ID=$TeamsDeepLinkEntityId",
  "DRY_RUN=$($DryRun.ToString().ToLower())"
)
Invoke-Az functionapp config appsettings set -n $FunctionAppName -g $ResourceGroupName --settings @settings | Out-Null
$current = Invoke-Az functionapp config appsettings list -n $FunctionAppName -g $ResourceGroupName
if ($current.name -contains 'AzureWebJobsStorage') {
  Invoke-Az functionapp config appsettings delete -n $FunctionAppName -g $ResourceGroupName --setting-names AzureWebJobsStorage | Out-Null
  Write-Host '   removed connection-string setting AzureWebJobsStorage (identity-based access is used instead)'
}
Write-Host "   $($settings.Count) settings set, DRY_RUN=$($DryRun.ToString().ToLower())"

Write-Host ''
Write-Host '== Done. Values for the next steps:' -ForegroundColor Green
[pscustomobject]@{
  FunctionAppName     = $FunctionAppName
  ResourceGroupName   = $ResourceGroupName
  IdentityClientId    = $identity.clientId
  IdentityPrincipalId = $identity.principalId
  StorageAccountName  = $StorageAccountName
} | Format-List
Write-Host 'Next:'
Write-Host "  1. .\Grant-NotificationPermissions.ps1 -ServicePrincipalObjectId $($identity.principalId)"
Write-Host "  2. Grant-PnPAzureADAppSitePermission -AppId $($identity.clientId) -DisplayName '$FunctionAppName' -Site <site> -Permissions Write   (per site, see README)"
Write-Host "  3. .\New-NotificationTeamsApp.ps1 -NotificationClientId $($identity.clientId) -SiteUrl $ConfigSiteUrl -AppName <name> -DeveloperName <org>"
if (-not $TeamsDeepLinkAppId) {
  Write-Host "     then: az functionapp config appsettings set -g $ResourceGroupName -n $FunctionAppName --settings TEAMS_DEEPLINK_APP_ID=<Teams app id from step 3>"
}
Write-Host "  4. .\Publish-NotificationFunction.ps1 -ResourceGroupName $ResourceGroupName -FunctionAppName $FunctionAppName"
