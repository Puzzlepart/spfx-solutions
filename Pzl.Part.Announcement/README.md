# Announcement (Driftsmeldinger)

[![version](https://img.shields.io/badge/version-1.5.1-orange.svg)](https://semver.org)

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
  * Column added in 1.5.0 (banner):
    * `PzlShowAsBanner` (Yes/No, default No): the announcement is also shown as a banner at the top of the pages by the **AnnouncementBanner** application customizer, see [Banner](#banner-application-customizer). Not read by the web part.

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

### Deep link redirect (application customizer)

The package also contains the application customizer **AnnouncementDeepLink** (component ID `9f3d6a2c-7b14-4e8f-a5c0-3d2e1f6b8a47`). It reads the `driftsmelding` query parameter when the page loads and, for `page:<url>` deep links, sends the user to that page right away. The web part does the same when it renders, but only if the section holding it is rendered: on small screens the right column of the home page is loaded late or not at all, so the redirect is left to the customizer, which runs before the page content regardless of layout. Only same-origin `.aspx` paths are accepted, and other query parameters (such as `app=portals` from the Teams tab) are carried over. `driftsmelding:<id>` deep links are still handled by the web part, which owns the popover.

The customizer is registered as a site custom action by the feature **Announcements by Puzzlepart - deep link redirect** when the package is installed on a site (site collection app catalog, or "add app" on the site). With a tenant-wide deployment from the tenant app catalog the feature is not activated per site, so register the custom action on the site the Teams tab opens (normally the root site) with the template:

```powershell
Connect-PnPOnline -Url "https://<tenant>.sharepoint.com" -Interactive -ClientId "<clientid>"
Invoke-PnPSiteTemplate -Path ".\Templates\DeepLinkCustomAction.xml"
```

or directly with `Add-PnPCustomAction -Name AnnouncementDeepLink -Title AnnouncementDeepLink -Location ClientSideExtension.ApplicationCustomizer -ClientSideComponentId 9f3d6a2c-7b14-4e8f-a5c0-3d2e1f6b8a47 -ClientSideComponentProperties "{}" -Scope Site`. Check with `Get-PnPCustomAction -Scope Site`. The customizer does nothing on pages opened without the `driftsmelding` parameter.

Lists provisioned with the 1.0.x template still work: the web part falls back to the original fields if `PzlPinned`/`PzlNotifiedAt`/`PzlShowAsBanner` are missing. Re-run `Invoke-PnPSiteTemplate` to add the new columns.

## Banner (application customizer)

The package also contains the application customizer **AnnouncementBanner** (component ID `c0d1edf4-db39-4feb-8052-a4c27ffe1954`). On every page of a site where it is registered it reads the announcement list and shows the active announcements with **Vis som banner** (`PzlShowAsBanner`) set as message bars in the page's top placeholder:

* Colour and icon follow the announcement's level (`PzlSeverity`), like in the web part. Pinned announcements come first and get the pin icon.
* The banner shows the title and the first 160 characters of the content. Clicking it (or Enter/Space) opens a dialog with the same content as the web part's popover: the fields in `popoverFields` (default: the web part's default popover fields) and the "notified in Teams" footer.
* The close button dismisses the banner for the user. The dismissal is kept in `localStorage` (`pzl-announcement-banner-dismissed`) per announcement and version (`Modified`), so an edited announcement comes back. Dismissals for announcements that are no longer active are pruned.
* Audience targeting is honoured when `targetAudience` is set, with the same rules as the web part.
* The customizer decides from the page context alone whether the page is targeted (see below). Only on targeted pages is the list code loaded and the list read, and the UI is loaded only when there is a banner to show. Pages outside the targeting cost a few KB of script, which is what makes the tenant-wide registration affordable.

### Where the banner is shown (targeting)

The banner is shown on a page when one of these properties matches. With none of them set, it is shown nowhere.

| Property | Matches |
| --- | --- |
| `siteUrls` | Absolute site URLs. A site collection URL also covers its subsites (`/sites/x` covers `/sites/x/sub`, not `/sites/xy`). The root site collection URL matches the root site and its subsites only. Case and trailing slash are ignored. |
| `hubSiteUrls` | The hub site's absolute URL or its hub site ID (GUID). Every site associated with that hub shows the banner, the hub site included. Sites are matched on their own hub only, so a hub associated to another hub must be listed itself. |
| `everywhere` | `true` shows the banner on every site the customizer is registered on. |

### Registering the banner

**Tenant-wide (hub and lists of sites).** When the package is deployed tenant-wide ("make available to all sites"), the feature adds a row for the customizer to the **Tenant Wide Extensions** list in the tenant app catalog, with no targeting. Edit the row (Site contents of the app catalog, column *Component Properties*) and set the list site and the targeting, e.g.

```json
{
  "listSiteUrl": "https://<tenant>.sharepoint.com",
  "siteUrls": ["https://<tenant>.sharepoint.com", "https://<tenant>.sharepoint.com/sites/x"],
  "hubSiteUrls": ["https://<tenant>.sharepoint.com/sites/abc"],
  "targetAudience": false
}
```

The change applies on the next page load; nothing is deployed or scripted per site. Sites that join the hub get the banner, sites that leave lose it. With PnP: `Set-PnPListItem -List "Tenant Wide Extensions" -Identity <id> -Values @{ TenantWideExtensionComponentProperties = '<json>' }` on the app catalog site. Was the package not deployed tenant-wide, the row can be added by hand with *Component Id* `c0d1edf4-db39-4feb-8052-a4c27ffe1954` and *Location* `ClientSideExtension.ApplicationCustomizer`.

**One site.** `Templates/BannerCustomAction.xml` registers the customizer on a single site collection and passes that site's URL as `siteUrls`, so a site owner can turn the banner on without touching the app catalog. `ListSiteUrl` points it at the list on the intranet site when registered elsewhere:

```powershell
Connect-PnPOnline -Url "https://<tenant>.sharepoint.com/sites/<site>" -Interactive -ClientId "<clientid>"
Invoke-PnPSiteTemplate -Path ".\Templates\BannerCustomAction.xml" -Parameters @{ ListSiteUrl = "https://<tenant>.sharepoint.com" }
```

A site covered by both registrations shows the banner once: the first instance on the page sets a marker on `window`, and the second does nothing. Remove a site registration with `Remove-PnPCustomAction -Identity AnnouncementBanner -Scope Site -Force`.

| Parameter / property | Default | Description |
| --- | --- | --- |
| `ListSiteUrl` / `listSiteUrl` | Current site | Site that hosts the Driftsmeldinger list. Users on the site must have read access to the list. Set it on the tenant-wide row. |
| `ListTitle` / `listTitle` | `Driftsmeldinger` | Title of the list |
| `TargetAudience` / `targetAudience` | `false` | Filter on the list's audience targeting column |
| `siteUrls`, `hubSiteUrls`, `everywhere` | none | Targeting, see above. The site template sets `siteUrls` to the site itself. |
| `popoverFields` (property only) | Web part defaults | Fields shown in the dialog, keyed by internal name, e.g. `{"PzlContent":true,"PzlEndDate":true}`. |

Application customizers are not rendered in the Viva Connections mobile app, and the web part keeps showing the announcement in its list regardless of the banner.

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

### 1.5.1

* Banner targeting: new properties `siteUrls`, `hubSiteUrls` and `everywhere` decide where the banner is shown, from the page context alone. With no targeting the banner is shown nowhere.
* Tenant-wide registration: the banner feature now carries a `ClientSideInstance.xml`, so a tenant-wide deployment adds the customizer to the Tenant Wide Extensions list with no targeting. Set `listSiteUrl` and the targeting on that row to show the banner on a hub and on chosen sites, with no per-site scripting; joining and leaving the hub is picked up on the next page load.
* `Templates/BannerCustomAction.xml` passes the site's own URL as `siteUrls`. A site covered by both a site registration and the tenant-wide row shows the banner once.
* The customizer's entry bundle holds only the targeting check; the list code is loaded on targeted pages and the UI when there is a banner to show.

### 1.5.0

* New column **Vis som banner** (`PzlShowAsBanner`) in the template and in the web part's form and popover toggles. Re-run `Templates/Announcement.xml` to add it.
* New application customizer **AnnouncementBanner** in the same package: shows the active announcements with **Vis som banner** set as message bars at the top of the pages, coloured by level, dismissable per user (remembered in `localStorage` until the announcement is edited), and with a dialog showing the same details as the web part's popover on click. Registered per site with `Templates/BannerCustomAction.xml`, which can point the banner at the list on another site (`ListSiteUrl`), so the whole hub can show the intranet's banners.
* The announcement fetch (active items, audience targeting, severity mapping) is shared between the web part and the banner (`services/fetchAnnouncements.ts`), and `AnnouncementListService` can read a list on another site.

### 1.4.1

* Fixed the 1.4.0 package, which the app catalog rejected with "Component ID ddfa2478-… exists more than once in the package": both features included every component. Each feature now lists its own component (`componentIds`), so the web part belongs to the web part feature and the customizer to the deep link feature. 1.4.0 never published its assets, so update to 1.4.1 before registering the custom action.

### 1.4.0

* New application customizer **AnnouncementDeepLink** in the same package: redirects `?driftsmelding=page:<url>` deep links (news notifications) to the news page as soon as the page's SPFx components load, without waiting for the web part. On small screens the column holding the web part was loaded late or not at all, so the redirect did not always happen. Registered by the new feature **Announcements by Puzzlepart - deep link redirect** on site install; `Templates/DeepLinkCustomAction.xml` registers it for tenant-wide deployments.
* The deep link parsing is shared between the web part and the customizer (`src/shared/deepLink.ts`). The web part keeps its redirect as a fallback for sites without the custom action.

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
