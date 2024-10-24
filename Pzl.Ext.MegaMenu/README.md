# Global navigation by Puzzlepart #

## Release notes ##

### 2.8.6.3 ###

Added an optional new hyperlink field "Link to more information" used by the ServiceAnnouncements list.

### 2.8.0.0 ###

Added an optional search bar to the navigation row, and moved most settings to a new sharepointlist named "NavSettings"

### 2.7.0.0 ###

Upgraded to Spfx 1.7.1. Improved mobile navigation by adding expand/collpase chevrons.

### 2.6.0.0 ###

Warning: Breaking change. Added the sort order field to the navigation links list as well, to allow for custom sort order of link elements (with fallback to alphabetical sorting like before). A consequence is that if you install the app, you need to ensure that the NavigationLinks list has the PzlNavOrder field. The easiest way to do this is to apply the ConfigurationLists.xml template to the cdn site. If this is not done, the menu will show an error.

## Installation guide ##

### Setting up a list data source ###

* Add a sitecollection that is called /sites/cdn (this can be configured, see guide below)
* Install the PnP package "ConfigurationLists.xml" from './templates/ConfigurationLists' with "Invoke-PnPSiteTemplate"
* Start populating the Nav Headings and Nav Links lists in the CDN site.

### Building the app-package ###

* Ensure npm and gulp is installed globally
* Navigate to the Pzl.MegaMenu folder and run `npm i`
* Ensure that the config files is set up correctly (env.json and serve.json if serving locally)
  * `env.json`: Should contain the tenant information as well as the correct tenant and folder in the cdn (see env.sample.json for guide)
  * `serve.json`: Should contain the tenant URL and feature properties for development and debugging (see serve.sample.json for guide)
  * `package-solution.json`: Mark both versions with the correct version number. Increment if you've done noticable changes.
  * `sharepoint/assets/elements.xml`: ClientSideComponentProperties needs to be marked with the correct server relative path to the cdn.
* Run one of the following
  * Dev and deploy: Run the deploy command: `npm run deploy` (automatic install, requires env.json to be filled out)
  * Package&Install: Run `gulp bundle --ship` and then `gulp package-solution --ship`. Finally upload the .sppkg app package from sharepoint/solution to the App Catalog in SharePoint Online.

### Adding the mega menu app ###

Navigate to the sitecollections where the megamenu should be rendered, go to site contents and select 'add an app'. Add 'Global Navigation by Puzzlepart'. After a few seconds the app should render.

## Changing the Navigation settings ##
After version 2.8.0.0 most of the settings (all except the list URL's) were moved to a Sharepoint list. This is installed by default when you install the the other necessary lists from the template folder, and is populated with some default values. Here you may edit the color of all the surfaces, texts and icons, as well as the option to add the following add-ins:
* Searchbar - A searchbar which redirects users to the site's search center with the entered query.
* HomeButton - A icon (could be anything, but the main intention was a home icon) with a link to the intended homepage of the tenant.
* HelpButton - A button (could be anything and point anywhere) which was originally made as a quicklink to a help centre, and is configured as such by default.

The settings list is created at /sites/cdn/lists/NavSettings and most fields should be self explanatory. A quick note on some of the datatypes:
* Color settings should be any color string accepted by css, as it is passed directly. [Definition](https://www.w3schools.com/cssref/pr_text_color.asp)
* Boolean settings (like "helpButtonEnabled") will only make the add-in appear if it contains the string "true" (without the quotation marks)
* searchBarSearchUrl - should be a link to the search center, but will NOT work if any query parameters (anyting followed by a: ?) already exist. 

Please make sure to NOT edit any of the setting titles, as they will stop working.

## Setting up a data source ##

The global navigation extension supports data from either a SP list or a Taxonomy Term Group. See config/package-solution.json and sharepoint/assets/elements.xml.

### SharePoint list as data source ###

Example configuration for a SP list:

```json
{
   "dataSource": {
        "spList": {
            "serverRelativeWebUrl": "/sites/cdn/",
            "linksListUrl": "Lists/NavLinks",
            "headersListLookupFieldName": "PzlNavLinkHeaderLookup",
            "urlFieldName": "PzlNavUrl",
            "headersOrderFieldName": "PzlNavOrder",
            "hasHeaderNavLinks": true
        }
    },
}
```

If you want to include ServiceAnnouncement-functionality aswell, add the following. The discardForSessionOnly controls the ability for the user to close the service announcement. If discardForSessionOnly is true, the notifications will return the next time the user opens the browser (difference between sessionStorage and localStorage).

```json
{
   "dataSource": {
        "spList": {
            "serverRelativeWebUrl": "/sites/cdn/",
            "linksListUrl": "Lists/NavLinks",
            "headersListLookupFieldName": "PzlNavLinkHeaderLookup",
            "urlFieldName": "PzlNavUrl",
            "headersOrderFieldName": "PzlNavOrder",
            "hasHeaderNavLinks": true
        }
    },
    "serviceAnnouncements": {
        "serverRelativeWebUrl": "/sites/cdn/",
        "listUrl": "Lists/serviceAnnouncements",
        "discardForSessionOnly": false,
        "textAlignment" : 1, /* 1 = Left, 2 = Center, 3 = Right */
        "boldText" : false
    },
}
```

### Taxonomy term group ###

Example configuration for term group with ID `e3174ce7-e0a9-4711-b9e8-c8c8bf3b7519` and expiration `1440` minutes (1 day):

**NB:** The termsets in the specified term group must have `Use this Term Set for Site Navigation` turned on.

```json
{
   "dataSource": {
        "taxonomy": {
            "termGroupId": "e3174ce7-e0a9-4711-b9e8-c8c8bf3b7519",
            "storageConfig": {
              "key": "pzl-global-navigation-data",
              "expirationMinutes": 1440
            }
        }
    },
}
```

## Configuration options ##

### Setting up a "Promoted Link" ###

A header at position -805 will gain the status "Promoted Link" and will appear *in* the navigationbar itself, with red background and a questionmark icon.
Only one link should be placed under this header, and it should only be used in its current state if the use is a support solution. 

### Setting up a "Home Link" ###

A header at position -127001 will gain the status "Home Link" and will appear *in* the navigationbar itself, with red background and a questionmark icon.
Only one link should be placed under this header, and it should only be used as a link to the starting page of the intranet

### Setting up custom colors for alert service announcement levels ###

By default service annoucements support four levels using standard colors

* Information / Informasjon
* Warning / Advarsel
* Alert / Varsel
* Normal / Normal

You may add additional levels which will use the Information icon, and by appending (#ff0000) to the choice, you can set custom colors for the levels.

Example:

* Information (red) <- will render Information type with a red background and information icon
* Alert (#00ffff) <- will render Alert type with a magenta background and alert icon

### Enabling CDN ###

We've seen that the solution does not always work if CDN is not enabled in the tenant. To enable CDN, do the following.

* Requires SharePoint Online Management Shell
* Connect-SPOService -Url https://tenantname-admin.sharepoint.com
* Set-SPOTenantCdnEnabled -CdnType Public

## Manual changing menu properties for one site using PnP Posh

It is possible to overide some of the default colors and texts in the mega menu. To to so use [PnP PowerShell](https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-pnp/sharepoint-pnp-cmdlets?view=sharepoint-ps) and log in as a site owner.

```ps
Connect-PnPOnline -Url https://innovationnorway.sharepoint.com/sites/intranet
$menu = Get-PnPCustomAction |? Title -eq "GlobalNavigation"
$properties = ConvertFrom-Json $menu.ClientSideComponentProperties
$properties

dataSource                 : @{spList=}
serviceAnnouncements       : @{serverRelativeWebUrl=/sites/intranet/; listUrl=Lists/serviceAnnouncements; discardForSessionOnly=False; textAlignment=2; boldText=True}
navToggleText              :
navToggleTextColor         :
navToggleBackgroundColor   :
navHeaderTextColor         :
navContentBackgroundColor  :
linkTextColor              :
supportLinkBackgroundColor :
supportLinkTextColor       :
navColumns                 : 5

# Change to use 4 columns, set the toggle text from **Menu** to **Something**, and the menu background color to blue.
$properties.navColumns = 4
$properties.navToggleText = "Something"
$properties.navToggleBackgroundColor = "#0000ff"
$menu.ClientSideComponentProperties = ConvertTo-Json $properties
$menu.Update()
Invoke-PnPQuery
```

## Example: Using PnP Modern Search to display Service Announcements
In this example we will demonstrate how you can use PnP Modern Search to display Service Announcements in the MegaMenu by Puzzlepart. Please follow the steps below.

1) Map the following managed properties on Tenant Level (https://tenantname-admin.sharepoint.com/_layouts/15/searchadmin/ta_listmanagedproperties.aspx?level=tenant). You can do this manually, or use PnP PowerShell to import the file located in /SearchConfiguration/TenantSearchConfiguration.xml.

| MANAGED PROPERTY NAME | MAPPED CRAWLED PROPERTIES | ALIASES |
| --- | ---| --- |
| RefinableDate10 | OWS_PZLSTARTDATE | PzlStartDate |
| RefinableDate11 | OWS_PZLENDDATE | PzlEndDate |
| RefinableString10 | OWS_PZLSEVERITY | PzlSeverity  |
| RefinableString11 | OWS_PZLAFFECTEDSYSTEMS | PzlAffectedSystems |
| RefinableString12 | OWS_PZLCONSEQUENCES | PzlConsequences |
| RefinableString13 | OWS_PZLRESPONSIBLE | PzlResponsible |

2) Install PnP Modern Search in your tenant - https://microsoft-search.github.io/pnp-modern-search/installation/

3) Upload the file /SearchConfiguration/PnPSearchResults_Template_Driftsmeldinger.html to a SharePoint site in your tenant, accessible to all employees, e.g.
    https://tenantname.sharepoint.com/sites/sitename/Shared%20Documents/PnPSearchTemplates.
4) Add the PnP Modern Search webpart to your site and configure it to use an external template URL (page 2 of 4 in the configuration of the web part). Set Available layouts (page 2 of 4) to "Custom".

The result afterwards should look like below.
 
![Service Announcements displayed using PnP Modern Search](/Pzl.Ext.MegaMenu/documentation/PnPSearchResults_Template_Driftsmeldinger.png)