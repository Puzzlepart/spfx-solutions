## pzl-ext-google-analytics

This is where you include your WebPart documentation.

### Building the code

```bash
git clone the repo
npm i
npm i -g gulp
gulp
```

This package produces the following:

* lib/* - intermediate-stage commonjs build artifacts
* dist/* - the bundled script, along with other resources
* deploy/* - all resources which should be uploaded to a CDN.

### Build options

gulp clean - TODO
gulp test - TODO
gulp serve - TODO
gulp bundle - TODO
gulp package-solution - TODO



Add-PnPCustomAction -ClientSideComponentId "9a936eb0-8370-418e-b4f7-3c6e7f952b50" -Name "GoogleAnalyticsExtension" -Title "GoogleAnalyticsExtension" -Location ClientSideExtension.ApplicationCustomizer -ClientSideComponentProperties: '{"trackerID":"<id>"}' -Scope site