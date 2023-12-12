# Connect to URL

$Url = "https://puzzlepart.sharepoint.com/sites/SharedResources"
Connect-PnPOnline -Url $Url -Interactive

Set-PnPTraceLog -On -LogFile "traceoutput.txt" -Level Debug

# Apply provisioning Template
$TemplateXML = ".\ConfigurationLists.xml"
Invoke-PnPSiteTemplate -path $TemplateXML