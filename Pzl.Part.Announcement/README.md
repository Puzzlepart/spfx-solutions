# Driftsmeldinger (Pzl.Part.Announcement)

[![version](https://img.shields.io/badge/version-2.0.0-orange.svg)](https://semver.org)

## Summary

Driftsmeldinger shows operational announcements (incidents, outages, planned maintenance) from one central SharePoint list, in a web part and optionally as a banner at the top of the pages. Each announcement has a **message level** (*Informasjon / Advarsel / Kritisk*) and a **status** (*Planlagt / Pågår / Løst*), and is shown between its start and end date. Audience targeting is supported. Users who may add items to the list get a **New announcement** button that opens a form built from the fields of the list; users who may edit items can change the status from the popover.

Teams notifications are an optional add-on, [Pzl.Ext.Notification](../Pzl.Ext.Notification/README.md). The terms used in the solution are defined in [CONTEXT.md](CONTEXT.md).

| Announcements                                                                             | Popover (on click)                                                                        |
| ----------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------- |
| ![image](https://github.com/user-attachments/assets/c8de6cc6-c843-4d54-957b-4d7e55b9e363) | ![image](https://github.com/user-attachments/assets/413b3a74-d682-4e88-b136-797c6fcbec7c) |

![image](https://github.com/user-attachments/assets/c934d56e-1590-4880-8877-8a545acb9cf6)

## The list

One list **Driftsmeldinger** per tenant, normally on the intranet's root site (`Templates/Announcement.xml`). The web part and the banner on other sites read it through **Område for driftsmeldinger** / `listSiteUrl`. Audience targeting must be turned on for the list if it is used.

| Column | Internal name | Description |
| --- | --- | --- |
| Tittel | `Title` | Required. |
| Meldingsnivå | `PzlSeverity` | Required, default *Advarsel*. *Informasjon*, *Advarsel* or *Kritisk*; sets the colour and icon. The 1.x values *Varsel* and *Normal* are shown as *Kritisk* and *Informasjon*. |
| Status | `PzlStatus` | Required, default *Planlagt*. *Planlagt*, *Pågår* or *Løst*. *Løst* is shown green whatever the level. Items without a status (created before 2.0) count as *Pågår*. |
| Innhold | `PzlContent` | Required, rich text. |
| Startdato, Sluttdato | `PzlStartDate`, `PzlEndDate` | Required. When the announcement is shown on the intranet, not when the incident or the work happens. |
| Systemer berørt | `PzlAffectedSystems` | |
| Eventuelle konsekvenser | `PzlConsequences` | Consequences that last after the issue is resolved. |
| Ansvarlig | `PzlResponsible` | Who users can contact. |
| Fest til toppen | `PzlPinned` | Pinned announcements come first, with a pin icon. |
| Vis som banner | `PzlShowAsBanner` | Also shown as a banner at the top of the pages, see [Banner](#banner-application-customizer). Preselected in the web part's form for *Kritisk*. |

Every column has a description in the list form that tells the editor when to use it. The Teams columns (`PzlNotifyTeams`, `PzlNotifyAgain`, `PzlNotifiedAt`) belong to the notification add-on and are added by its template `AnnouncementFields.xml`.

Announcements are sorted: pinned first, then status (*Pågår*, *Planlagt*, *Løst*), then level (*Kritisk* first), then the newest start date. Level and status are read with fixed tables (`src/Announcement/model.ts`) in Norwegian and English, independent of the user's language.

## Web part properties

| Group | Property | Description |
| --- | --- | --- |
| Generelt | Tittel, Beskrivelse | Header title and the text behind the info icon. |
| Generelt | Område for driftsmeldinger | Absolute URL of the site with the Driftsmeldinger list (the central list). Empty means the current site. Users need read access to the list; the create button and the quick actions check the user's permissions on that list. |
| Generelt | Aktiver målgruppeangivelse | Filter announcements on the list's audience targeting column. |
| Felter i detaljvisning | One toggle per list field | Which fields are shown in the popover (on click). The toggles are read from the default content type of the list, so columns added to the list show up here. Title is always shown. Defaults match 1.2.0: content, start/end date, affected systems, consequences, responsible and "notified in Teams". |
| Felter i skjema for ny driftsmelding | One toggle per editable list field | Which fields are shown in the form for new announcements. Required fields and fields hidden in the list cannot be changed here. «Varsle i Teams» (notification add-on) is off by default. |
| Skjul/vis | Skjul header | Hide the header (the button for new announcements is still shown). |
| Skjul/vis | Skjul knapp for ny driftsmelding | Hide the **New announcement** button for everyone. |

### Registering announcements from the web part

The **New announcement** button is shown only to users who have the *Add items* permission on the list (checked with `currentUserHasPermissions`), and only if it is not hidden in the web part properties. It opens a dialog with a form generated from the fields of the list's default content type (fields that are hidden in the column or in the content type, read-only or have `ShowInNewForm="FALSE"` on the site or list column are left out). Supported column types: single/multiple lines of text, choice (single and multi), date/time, yes/no, person or group (single and multi), number, currency and hyperlink. Other column types are not shown in the form and get their default value from the list. Required columns are validated before saving. Rich text columns (such as *Innhold*) get a small formatting toolbar (bold, italic, underline, lists, link) and are stored as HTML, like in the list form. «Vis som banner» follows the level until the user changes it: on for *Kritisk*, off otherwise. `PzlNotifiedAt` and `PzlNotifyAgain` (notification add-on) are never shown in the form.

### Quick actions

Users with the *Edit items* permission on the list get buttons at the bottom of the popover: **Marker som pågår** (planned announcements), **Marker som løst** (planned and ongoing) and **Rediger**, which opens the list's edit form and returns to the page. The status is written with the list's own choice value, so a list provisioned in English gets *Resolved*. The quick actions are shown only when the list has the *Status* column.

## Teams notifications and deep links

The web part itself does not send notifications. Teams notifications are an optional add-on, [Pzl.Ext.Notification](../Pzl.Ext.Notification/README.md) (Azure Function, a small Teams app and an SPFx package that redirects news notifications). The list and the web part work without it.

To open a specific announcement when the page loads, the web part reads an announcement ID from one of these sources (in this order) and opens its popover:

1. Query string: `?driftsmelding=<id>` on the page URL, so a link to one announcement can be shared. The value may carry the deep link prefix (`driftsmelding:<id>`).
2. Teams deep link: `subEntityId`/`subPageId` set to `driftsmelding:<id>` (read via `context.sdks.microsoftTeams`).

`page:<url>` values (news notifications) are ignored by the web part and handled by the add-on's **NotificationDeepLink** customizer.

Lists provisioned with an older template still work: the web part falls back to the original fields if `PzlPinned`/`PzlShowAsBanner`/`PzlStatus` are missing. Re-run `Invoke-PnPSiteTemplate` to add the new columns, see [Upgrading from 1.x](#upgrading-from-1x).

## Banner (application customizer)

The package also contains the application customizer **AnnouncementBanner** (component ID `c0d1edf4-db39-4feb-8052-a4c27ffe1954`). On every page of a site where it is registered it reads the announcement list and shows the active announcements with **Vis som banner** (`PzlShowAsBanner`) set as message bars in the page's top placeholder:

* Colour and icon follow the announcement's level and status, like in the web part: *Løst* is green with the label «Løst», *Planlagt* has a calendar icon and the label «Planlagt». Pinned announcements come first and get the pin icon.
* The banner shows the title and the first 160 characters of the content. Clicking it (or Enter/Space) opens a dialog with the same content as the web part's popover: the fields in `popoverFields` (default: the web part's default popover fields) and the "notified in Teams" footer.
* The close button dismisses the banner for the user. The dismissal is kept in `localStorage` (`pzl-announcement-banner-dismissed`) per announcement and version (`Modified`), so an edited announcement comes back. That is how users who dismissed a critical banner see it once more, green, when it is marked as resolved. Dismissals for announcements that are no longer active are pruned.
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

**Tenant-wide (hub and lists of sites).** The package must be deployed tenant-wide ("make available to all sites"). Then add one row for the customizer to the **Tenant Wide Extensions** list on the tenant app catalog site (Site contents), with the list site and the targeting in *Component Properties*:

```powershell
Connect-PnPOnline -Url "https://<tenant>.sharepoint.com/sites/appcatalog" -Interactive -ClientId "<clientid>"
Add-PnPListItem -List "Tenant Wide Extensions" -Values @{
  Title                                  = "AnnouncementBanner"
  TenantWideExtensionComponentId         = "c0d1edf4-db39-4feb-8052-a4c27ffe1954"
  TenantWideExtensionLocation            = "ClientSideExtension.ApplicationCustomizer"
  TenantWideExtensionComponentProperties = '{"listSiteUrl":"https://<tenant>.sharepoint.com","siteUrls":["https://<tenant>.sharepoint.com/sites/x"],"hubSiteUrls":["https://<tenant>.sharepoint.com/sites/abc"],"targetAudience":false}'
  TenantWideExtensionSequence            = 110
  TenantWideExtensionListTemplate        = 0
  TenantWideExtensionDisabled            = $false
}
```

The same can be done with **New item** in the list: Title, Component Id, Component Properties, Location `ClientSideExtension.ApplicationCustomizer`, Sequence 110, Disabled = No; leave Web template and List template empty. The row applies on the next page load and is edited in place later (`Set-PnPListItem -Identity <id> -Values @{ TenantWideExtensionComponentProperties = '<json>' }`); nothing is deployed or scripted per site. Sites that join the hub get the banner, sites that leave lose it.

The package does not create this row itself. SPFx can ship a `ClientSideInstance.xml` for that, but SharePoint processes it only on a first-time tenant-wide deployment (never on an upgrade), and in 1.5.1 a first-time deployment failed in a tenant where the row already existed, so it was removed again in 1.5.2.

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

### With the install script (new installation and upgrade)

`scripts/Install-Driftsmeldinger.ps1` does the steps below in one go and can be run again for every upgrade. It is idempotent and supports `-WhatIf`. It needs PnP.PowerShell (not installed by the script), SharePoint Administrator rights and the package built with `npm run package`.

```powershell
.\scripts\Install-Driftsmeldinger.ps1 -SiteUrl "https://<tenant>.sharepoint.com" -WhatIf
.\scripts\Install-Driftsmeldinger.ps1 -SiteUrl "https://<tenant>.sharepoint.com" -BannerHubSiteUrls "https://<tenant>.sharepoint.com/sites/<hub>"
```

`-ClientId` defaults to the Prosjektportalen (PP365) PnP app `da6c31a6-b557-4ac3-9994-7315da06ea3a`; pass `-ClientId` if the tenant uses its own app registration for PnP.PowerShell.

| Step | What it does |
| --- | --- |
| Package | Uploads and publishes `pzl-announcement.sppkg` to the tenant app catalog (replaces 1.x in place) and updates the app on the list site if it was added there. `-SkipPackage` skips it. |
| List | Applies `Templates/Announcement.xml` to `-SiteUrl`, with the workaround for PnP.PowerShell 3.4.x. Detects a new installation, an upgrade from 1.x (no *Status* column) or an update of 2.x. |
| Cleanup | Removes the 1.x `AnnouncementDeepLink` custom action from the list site and the notification site. |
| Banner | With `-BannerSiteUrls`, `-BannerHubSiteUrls` or `-BannerEverywhere` (and `-BannerTargetAudience`): creates or updates the banner's row in Tenant Wide Extensions. Properties not given, such as `popoverFields`, are kept. |
| Teams notifications | With `-EnableTeamsNotifications`: uploads `pzl-ext-notification.sppkg` and applies `AnnouncementFields.xml` (list site), `NotificationConfig.xml` and `DeepLinkCustomAction.xml` (`-NotificationSiteUrl`, default the list site). The Azure, Entra and Teams part is done by hand, see [Pzl.Ext.Notification](../Pzl.Ext.Notification/README.md). |

The script ends with the manual steps that are left, such as adding the web part to a page and setting **Område for driftsmeldinger** on other sites.

### By hand

#### Create the list on the site that hosts it (normally the root site)

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
> The same applies to the templates in [Pzl.Ext.Notification](../Pzl.Ext.Notification/README.md).

#### Upload the package to an app catalog

Upload `sharepoint/solution/pzl-announcement.sppkg` from the build to the tenant app catalog (or a site collection app catalog), add the web part to a page and set **Område for driftsmeldinger** on sites other than the one with the list.

### Upgrading from 1.x

2.0.0 keeps the solution ID and the component IDs, so it is an in-place upgrade: existing web part instances and the banner's Tenant Wide Extensions row keep working. `scripts/Install-Driftsmeldinger.ps1` does steps 1 to 3; the steps by hand:

1. Upload the 2.0.0 package to the app catalog where 1.x is deployed (replace the existing app). Sites with the app installed from a site collection app catalog need **Update** under Site contents.
2. Re-run `Invoke-PnPSiteTemplate -Path ".\Templates\Announcement.xml"` on the site with the list. This adds *Status*, changes the level choices to *Informasjon / Advarsel / Kritisk* and updates the column descriptions. Existing items keep their values: *Varsel* is shown as *Kritisk*, *Normal* as *Informasjon*, and items without a status as *Pågår*. When such an item is edited, the editor picks a new level and a status.
3. If the 1.4 to 1.5 deep link customizer (**AnnouncementDeepLink**) was registered on a site, remove it, since the component is no longer in the package: `Remove-PnPCustomAction -Identity AnnouncementDeepLink -Scope Site -Force`. Its replacement, **NotificationDeepLink**, is part of [Pzl.Ext.Notification](../Pzl.Ext.Notification/README.md) and only needed with Teams notifications.
4. The columns *Varsle i Teams* and *Varslet i Teams* stay on lists that had them (the template never deletes columns). They do nothing without the notification add-on and are hidden in the web part's form by default; delete them in the list settings if they are not wanted.
5. On web parts placed on other sites than the list, set **Område for driftsmeldinger**.

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

### Unit tests

```bash
npm test
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

### 2.0.0

* New name and publisher: **Driftsmeldinger** by SoftwareOne. The solution ID and the component IDs are unchanged.
* Teams notifications are a separate add-on, [Pzl.Ext.Notification](../Pzl.Ext.Notification/README.md): the notification service, its Teams app, the configuration list template and the deep link customizer (now **NotificationDeepLink**, new component ID) moved there. The Teams columns moved from the main template to the add-on's `AnnouncementFields.xml`. The web part no longer handles `page:` deep links.
* New column **Status** (`PzlStatus`): *Planlagt / Pågår / Løst*, required, default *Planlagt*. *Løst* is shown green with a check mark, *Planlagt* with a calendar icon, in the web part, the popover and the banner.
* Message levels renamed to *Informasjon / Advarsel / Kritisk* («Varsel» was easy to confuse with Teams notifications; «Normal» is removed). Levels and status are read with fixed tables in Norwegian and English, not with the user interface strings.
* Column descriptions that tell editors when to use each column.
* Sorting: pinned, status, level, newest start date.
* Quick actions in the popover for users who may edit the list: **Marker som pågår**, **Marker som løst**, **Rediger**.
* New web part property **Område for driftsmeldinger** (`listSiteUrl`): the web part reads the central list on another site, like the banner.
* «Vis som banner» is preselected in the web part's form when the level is *Kritisk*.
* «Varsle i Teams» is hidden in the web part's form unless turned on.
* Unit tests for the level and status mapping and the sorting (`npm test`).
* New install and upgrade script `scripts/Install-Driftsmeldinger.ps1`.

### 1.5.4

* Fixed the deployment failure of 1.5.1 to 1.5.3 on a first-time deployment ("Deployment failed" with only a correlation ID): the description of the banner feature was longer than 255 characters, which the app catalog's component manifest list cannot store, so the banner component was never registered. Feature descriptions are now short. Keep feature titles and descriptions under 255 characters.

### 1.5.3

* New solution ID (`cd056a9e-7e02-48ae-890d-bf87e815febd`) and new feature IDs. In a tenant where the previous package had been deleted from the tenant app catalog, a fresh deployment of the old solution ID kept failing with "Deployment failed" and a correlation ID, while the same package deployed fine to a site collection app catalog; the tenant catalog keeps state for a deleted solution ID that emptying the recycle bins does not clear. The component IDs are unchanged, so existing web part instances, the deep link custom action and the Tenant Wide Extensions row keep working. In the app catalog this is a new app: deploy 1.5.3, then delete the old "Announcements by Puzzlepart" row (solution ID `a65ac2f2-…`) if it is still there. The assets get a new CDN folder.

### 1.5.2

* Removed `ClientSideInstance.xml` from the banner feature. SharePoint only processes it on a first-time tenant-wide deployment (an upgrade never creates the row), and a first-time deployment of 1.5.1 failed in a tenant where the row had been added by hand. The Tenant Wide Extensions row is created by hand or with PnP, see [Registering the banner](#registering-the-banner).

### 1.5.1

* Banner targeting: new properties `siteUrls`, `hubSiteUrls` and `everywhere` decide where the banner is shown, from the page context alone. With no targeting the banner is shown nowhere.
* Tenant-wide registration: a row for the customizer in the Tenant Wide Extensions list, with `listSiteUrl` and the targeting, shows the banner on a hub and on chosen sites with no per-site scripting; joining and leaving the hub is picked up on the next page load. (1.5.1 shipped a `ClientSideInstance.xml` to create the row; removed in 1.5.2, see below.)
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
