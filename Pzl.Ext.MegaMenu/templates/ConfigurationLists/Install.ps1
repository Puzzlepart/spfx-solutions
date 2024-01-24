# Description: This script will install the configuration lists for the Mega Menu solution
# Usage: Run the script from a PowerShell console
# Example: .\Install.ps1 -Url https://contoso.sharepoint.com/sites/megamenuconfig
param(
    [Parameter(Mandatory=$true)]
    [string]$Url,
    [Parameter(Mandatory=$false)]
    [string]$TraceLogFile = "traceoutput.txt",
    [Parameter(Mandatory=$false)]
    [string]$TraceLevel = "Debug"
)

$TemplateXML = ".\ConfigurationLists.xml"

# Connect to SharePoint Online
Connect-PnPOnline -Url $Url -Interactive
Set-PnPTraceLog -On -LogFile $TraceLogFile -Level $TraceLevel

# Apply provisioning Template
Invoke-PnPSiteTemplate -path $TemplateXML