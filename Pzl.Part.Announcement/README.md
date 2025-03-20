# Announcement (Driftsmeldinger)

[![version](https://img.shields.io/badge/version-1.0.0-orange.svg)](https://semver.org)

## Summary

This is a web part that displays announcements from a SharePoint list. The announcements are displayed based on startdate and enddate. The web part support Target Audience for the announcements added to the list.

[picture of the solution in action, if possible]

## Lists

* Driftsmeldinger (Announcement)
  * Entries for announcements.
  * Target Audience is activated for the list and can be used if needed.

## Installation

### Create the needed lists on the site where you want to host the quick links solutions

Clone the project or download all artefacts. The template `Announcement.xml` is located in the Templates-folder. Use PnP.PowerShell 1.12 or later to install, see example:

```powershell
Connect-PnPOnline -Url "https://<tenant>.sharepoint.com/sites/<site>" -Interactive -ClientId "<clientid>" 
Invoke-PnPSiteTemplate -Path ".\Templates\Announcement.xml"
```

### Upload the web part package to a site collection app catalog

This can be done manually by navigating to the app catalog and uploading the .sppkg package from the build.

## Building

### Building the code

```bash
git clone the repo
npm i
npm run build
```

### Testing

You can test/debug using
`npm run serve`

```html
https://<tenant>.sharepoint.com/sites/<site>?debug=true&noredir=true&debugManifestsFile=https://localhost:4321/temp/manifests.js
```

### Building the code for production

```bash
npm run package
```
