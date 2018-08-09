## Google analytics SPFx extension

This webpart makes it possible to gather analytics on modern SharePoint sites with Google Analytics and Google Tag Manager.

### Building the code

```bash
git clone https://github.com/Puzzlepart/spfx-solutions.git
npm i
gulp --ship
gulp package-solution --ship
```

###

### Installing
* Copy `pzl-ext-google-analytics.sppkg` from `sharepoint\solution` and install it in your tenant.
* Run the following PowerShell cmdlets
    * `Connect-PnPOnline https://tenant.sharepoint.com/sites/site`
    * `Add-PnPCustomAction -ClientSideComponentId "9a936eb0-8370-418e-b4f7-3c6e7f952b50" -Name "GoogleAnalyticsExtension" -Title "GoogleAnalyticsExtension" -Location ClientSideExtension.ApplicationCustomizer -ClientSideComponentProperties: '{"trackerID":"GTM-XXXXXXX"}' -Scope site`
* Enjoy the stats!

