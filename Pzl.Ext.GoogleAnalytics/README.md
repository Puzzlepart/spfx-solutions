## Google analytics SPFx extension

This webpart makes it possible to gather analytics on modern SharePoint sites with Google Analytics and Google Tag Manager. By default it is a Tenant Wide Deploy.

### Get started

#### Clone the repo 
```bash
git clone https://github.com/Puzzlepart/spfx-solutions.git
```

#### Pre-build
Before building the code remember to update the *trackerID* property in the following files
* sharepoint/assets/ClientSideInstance.xml
* sharepoint/elements.xml

Update config/package-solution.json if you don't want to do a tenant wide deploy

#### Build the code
```bash
npm i
gulp --ship
gulp package-solution --ship
```

###

### Installing
* Copy `pzl-ext-google-analytics.sppkg` from `sharepoint\solution` and install it in your tenant.
* Either deploy tenant wide (default) or run the following PowerShell cmdlets (requires changes in config/package-solution.json)
    * `Connect-PnPOnline https://tenant.sharepoint.com/sites/site`
    * `Add-PnPCustomAction -ClientSideComponentId "9a936eb0-8370-418e-b4f7-3c6e7f952b50" -Name "GoogleAnalyticsExtension" -Title "GoogleAnalyticsExtension" -Location ClientSideExtension.ApplicationCustomizer -ClientSideComponentProperties: '{"trackerID":"GTM-XXXXXXX"}' -Scope site`
* Enjoy the stats!

