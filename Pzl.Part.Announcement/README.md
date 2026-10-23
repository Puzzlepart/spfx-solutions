# Announcement (Driftsmeldinger)

[![version](https://img.shields.io/badge/version-1.1.0-orange.svg)](https://semver.org)

## Summary

This is a web part that displays announcements from a SharePoint list. The announcements are displayed based on startdate and enddate. The web part support Target Audience for the announcements added to the list.

| Announcements                                                                             | Popover (on click)                                                                        |
| ----------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------- |
| ![image](https://github.com/user-attachments/assets/c8de6cc6-c843-4d54-957b-4d7e55b9e363) | ![image](https://github.com/user-attachments/assets/413b3a74-d682-4e88-b136-797c6fcbec7c) |

![image](https://github.com/user-attachments/assets/c934d56e-1590-4880-8877-8a545acb9cf6)

## Lists

* Driftsmeldinger (Announcement)
  * Entries for announcements.
  * Target Audience needs to be activated for the list and can be used if needed.
  * Columns added in 1.1.0 (Teams notification):
    * `PzlPinned` (Yes/No): pinned announcements are sorted first and shown with a pin icon.
    * `PzlNotifyTeams` (Yes/No, default No): ask the notification service to send a Teams activity notification for levels other than "Varsel". Not read by the web part.
    * `PzlNotifiedAt` (Date/Time, hidden in forms): set by the notification service when the notification has been sent. Shown in the popover as "Varslet i Teams dd.mm. kl. hh:mm".

## Teams notifications and deep links

The web part itself does not send notifications. A separate service (for example an Azure Function) reads the list, sends Teams activity notifications and writes `PzlNotifiedAt`. See the customer-specific setup for details.

To open a specific announcement when the page loads, the web part reads an announcement ID from one of these sources (in this order) and opens its popover:

1. Teams deep link: `subEntityId`/`subPageId` set to `driftsmelding:<id>` (read via `context.sdks.microsoftTeams`).
2. Query string: `?driftsmelding=<id>` on the page URL.

Lists provisioned with the 1.0.x template still work: the web part falls back to the original fields if `PzlPinned`/`PzlNotifiedAt` are missing. Re-run `Invoke-PnPSiteTemplate` to add the new columns.

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

## Changelog

### 1.1.0

* Pinned announcements (`PzlPinned`) are sorted first and shown with a pin icon.
* Opens the popover for an announcement given by a Teams deep link (`driftsmelding:<id>`) or `?driftsmelding=<id>`.
* Shows when the announcement was notified in Teams (`PzlNotifiedAt`).
* Template: new columns `PzlPinned`, `PzlNotifyTeams` and `PzlNotifiedAt`.

### 1.0.0

* Initial release.
