# Google analytics SPFx extension

This webpart makes it possible to gather analytics on modern SharePoint sites with Google Analytics and Google Tag Manager. By default it is a Tenant Wide Deploy.

## Get started

### Clone the repo

```bash
git clone https://github.com/Puzzlepart/spfx-solutions.git
```

### Pre-build

Update config/package-solution.json if you don't want to do a tenant wide deploy

### Build the code

```bash
npm i
gulp --ship
gulp package-solution --ship
```

## Installing

* Copy `pzl-ext-google-analytics.sppkg` from `sharepoint\solution` and install it in your tenant and deploy
* Deploy app
* Update Tracker ID in Component Properties in the list 'Tenant wide deploy', e.g. https://pzlcloud.sharepoint.com/sites/appcatalog/Lists/TenantWideExtensions/AllItems.aspx

### Installing to selected sites

* Add ap to site XXX
* Run the following PowerShell cmdlets (requires changes in config/package-solution.json). Update trackerID below.
  * `Connect-PnPOnline https://tenant.sharepoint.com/sites/XXX`
  * `Add-PnPCustomAction -ClientSideComponentId "9a936eb0-8370-418e-b4f7-3c6e7f952b50" -Name "GoogleAnalyticsExtension" -Title "GoogleAnalyticsExtension" -Location ClientSideExtension.ApplicationCustomizer -ClientSideComponentProperties: '{"trackerID":"GTM-XXXXXXX"}' -Scope site`

Enjoy the stats!
