# Announcement (Driftsmeldinger)

[![version](https://img.shields.io/badge/version-1.3.3-orange.svg)](https://semver.org)

## Summary

This is a web part that displays announcements from a SharePoint list. The announcements are displayed based on startdate and enddate. The web part support Target Audience for the announcements added to the list. Users who may add items to the list get a **New announcement** button that opens a form built from the fields of the list, and the fields shown in the popover are selected in the web part properties.

| Announcements                                                                             | Popover (on click)                                                                        |
| ----------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------- |
| ![image](https://github.com/user-attachments/assets/c8de6cc6-c843-4d54-957b-4d7e55b9e363) | ![image](https://github.com/user-attachments/assets/413b3a74-d682-4e88-b136-797c6fcbec7c) |

![image](https://github.com/user-attachments/assets/c934d56e-1590-4880-8877-8a545acb9cf6)

## Lists

* Driftsmeldinger (Announcement)
  * Entries for announcements.
* Varslingskonfigurasjon (optional, `Templates/NotificationConfig.xml`)
  * One row per site for the Teams notification service, see [notification/README.md](notification/README.md).
  * Target Audience needs to be activated for the list and can be used if needed.
  * Columns added in 1.1.0 (Teams notification):
    * `PzlPinned` (Yes/No): pinned announcements are sorted first and shown with a pin icon.
    * `PzlNotifyTeams` (Yes/No, default No): ask the notification service to send a Teams activity notification for levels other than "Varsel". Not read by the web part.
    * `PzlNotifiedAt` (Date/Time, hidden in forms): set by the notification service when the notification has been sent. Shown in the popover as "Varslet i Teams dd.mm. kl. hh:mm".

## Web part properties

| Group | Property | Description |
| --- | --- | --- |
| Generelt | Tittel, Beskrivelse | Header title and the text behind the info icon. |
| Generelt | Aktiver målgruppeangivelse | Filter announcements on the list's audience targeting column. |
| Felter i detaljvisning | One toggle per list field | Which fields are shown in the popover (on click). The toggles are read from the default content type of the list, so columns added to the list show up here. Title is always shown. Defaults match 1.2.0: content, start/end date, affected systems, consequences, responsible and "notified in Teams". |
| Felter i skjema for ny driftsmelding | One toggle per editable list field | Which fields are shown in the form for new announcements. Required fields and fields hidden in the list cannot be changed here. |
| Skjul/vis | Skjul header | Hide the header (the button for new announcements is still shown). |
| Skjul/vis | Skjul knapp for ny driftsmelding | Hide the **New announcement** button for everyone. |

### Registering announcements from the web part

The **New announcement** button is shown only to users who have the *Add items* permission on the list (checked with `currentUserHasPermissions`), and only if it is not hidden in the web part properties. It opens a dialog with a form generated from the fields of the list's default content type (fields that are hidden in the column or in the content type, read-only or have `ShowInNewForm="FALSE"` on the site or list column are left out). Supported column types: single/multiple lines of text, choice (single and multi), date/time, yes/no, person or group (single and multi), number, currency and hyperlink. Other column types are not shown in the form and get their default value from the list. Required columns are validated before saving. Rich text columns (such as *Innhold*) get a small formatting toolbar (bold, italic, underline, lists, link) and are stored as HTML, like in the list form. `PzlNotifiedAt` is never shown in the form since it is written by the notification service.

## Teams notifications and deep links

The web part itself does not send notifications. The optional notification service in [`notification/`](notification/README.md) (Azure Function + a small Teams app) reads the lists, sends Teams activity notifications for new news pages and active announcements, and writes `PzlNotifiedAt`. The list and the web part work without it; see [notification/README.md](notification/README.md) for the deployment guide.

To open a specific announcement when the page loads, the web part reads an announcement ID from one of these sources (in this order) and opens its popover:

1. Teams deep link: `subEntityId`/`subPageId` set to `driftsmelding:<id>` (read via `context.sdks.microsoftTeams`).
2. Query string: `?driftsmelding=<id>` on the page URL. The value may carry the deep link prefix (`driftsmelding:<id>`); `page:<server-relative url>` navigates to that page instead (news notifications).

Lists provisioned with the 1.0.x template still work: the web part falls back to the original fields if `PzlPinned`/`PzlNotifiedAt` are missing. Re-run `Invoke-PnPSiteTemplate` to add the new columns.

## Installation

### Create the needed lists on the site where you want to host the quick links solutions

Clone the project or download all artefacts. The template `Announcement.xml` is located in the Templates-folder. Use PnP.PowerShell 1.12 or later to install, see example:

```powershell
Connect-PnPOnline -Url "https://<tenant>.sharepoint.com/sites/<site>" -Interactive -ClientId "<clientid>" 
Invoke-PnPSiteTemplate -Path ".\Templates\Announcement.xml"
```

> **Known issue with PnP.PowerShell 3.4.x.** `Invoke-PnPSiteTemplate` fails with `Could not load file or assembly 'Microsoft.SharePoint.Client.UserProfiles, Version=16.1.0.0'` on PnP.PowerShell 3.4.0, 3.4.1 and the current nightly. The assembly resolver introduced in 3.4.0 does not resolve CSOM assemblies that are not loaded at import, and the templates in this solution use localization (`<pnp:Localizations>`), which makes PnP.Framework load `UserProfiles`. Either use PnP.PowerShell 3.2.0 (verified) or pre-load the CSOM assemblies after importing the module and before connecting:
>
> ```powershell
> Import-Module PnP.PowerShell
> Get-ChildItem (Join-Path (Get-Module PnP.PowerShell).ModuleBase 'Core') -Filter 'Microsoft.*SharePoint*.dll' |
>   ForEach-Object { [System.Reflection.Assembly]::LoadFrom($_.FullName) | Out-Null }
> ```
>
> The same applies to `Templates\NotificationConfig.xml`.

### Upload the web part package to a site collection app catalog

This can be done manually by navigating to the app catalog and uploading the .sppkg package from the build.

## Building

Requires Node 22 (`>=22.14.0 <23`). The project uses SPFx 1.23 with the Heft toolchain (no gulp).

### Building the code

```bash
git clone the repo
npm i
npm run build
```

### Testing

You can test/debug using `npm run serve` (runs `heft start`), then open

```html
https://<tenant>.sharepoint.com/sites/<site>?debug=true&noredir=true&debugManifestsFile=https://localhost:4321/temp/build/manifests.js
```

### Linting

```bash
npm run lint
```

### Building the code for production

```bash
npm run package
```

The package is written to `sharepoint/solution/pzl-announcement.sppkg`.

## Changelog

### 1.3.3

* News deep links: when the page is opened with `?driftsmelding=page:<server-relative page url>` (what the notification Teams app's tab passes for news notifications), the web part navigates to that page. Only same-origin `.aspx` paths are accepted.

### 1.3.2

* Deep link focus: the `driftsmelding` query parameter is read before the Teams context, and the Teams SDK call has a 3 s timeout. The notification Teams app passes `{subEntityId}` in the tab URL, so focus works in Teams tabs without relying on `subPageId`.

### 1.3.1

* Fixed reading of `ShowInNewForm`, `RichText`, `Format` and `UserSelectionMode` from the list schema (the attributes were never matched), so fields hidden from the new form are left out and rich text columns are detected. Typed field properties (`RichText`, `DisplayFormat`, `SelectionMode`) and the content type's field links are used too.
* Popover renders rich text (HTML) columns as HTML.
* Rich text columns get a formatting toolbar in the form for new announcements and are stored as HTML.
* New property group **Felter i skjema for ny driftsmelding** to hide fields from the form.

### 1.3.0

* Web part property group **Felter i detaljvisning**: choose which list fields are shown in the popover. The fields are read from the list, so custom columns can be shown too.
* **New announcement** button with a dialog that creates list items. The form is generated from the fields of the list. Shown only to users with *Add items* permission on the list, and can be hidden with **Skjul knapp for ny driftsmelding**.
* Popover labels now use the display names of the list columns. Date columns are shown with time only if the column format is *Date & Time*.

### 1.2.0

* Upgraded to SPFx 1.23.2 (Heft toolchain, TypeScript 5.8, Node 22). gulp, `gulpfile.js` and `.eslintrc.yaml` are replaced by Heft and `eslint.config.js`.
* Updated Fluent UI, PnPjs, ESLint and Prettier to latest stable.
* Pinned `tabster` to 8.8.0 to match the version SharePoint ships (avoids a crash in SharePoint's own Fluent UI v9 components on first load).

### 1.1.0

* Pinned announcements (`PzlPinned`) are sorted first and shown with a pin icon.
* Opens the popover for an announcement given by a Teams deep link (`driftsmelding:<id>`) or `?driftsmelding=<id>`.
* Shows when the announcement was notified in Teams (`PzlNotifiedAt`).
* Template: new columns `PzlPinned`, `PzlNotifyTeams` and `PzlNotifiedAt`.

### 1.0.0

* Initial release.
