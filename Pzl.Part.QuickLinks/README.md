# Quick Links by Puzzlepart

Solution consists of two web parts and two lists.

* Quick links web part - add to a page for managing all links
* All links web part - used for adding links for the user

## Web parts

### Quick links web part - add to a page for managing all links

![Quick links web part preview](./preview/quick_links_webpart_06-12-19png.png "Web part")
![Quick links property pane preview](./preview/quick_links_property_pane_webpart_06-12-19png.png "Property pane")

### All links web part - used for adding links for the user

![All links web part preview](./TODO_ADD_IMAGE "Web part")
![All links property pane preview](./TODO_ADD_IMAGE "Property pane")

## Lists

* EditorLinks
  * Entries for links. You can add both mandatory and optional links.
  * Icons can be names from Office UI Fabric - [https://developer.microsoft.com/en-us/fabric#/styles/icons](https://developer.microsoft.com/en-us/fabric#/styles/icons)
* FavouriteLinks
  * It's important that all employees have write access to this list
  * The list stores one entry per user

## Installation

### Create the needed lists on the site where you want to host the quick links solutions

```powershell
Connect-PnPOnline -Url https://tenantnamehere.sharepoint.com/sites/intranet
Apply-PnPProvisioningTemplate -Path .\templates\QuickLinks.xml
```

### Upload the web part package to a site collection app catalog

```powershell
Connect-PnPOnline -Url https://tenantnamehere.sharepoint.com/sites/intranet
# Create app catalog if not present
$site = Get-PnPSite
# Upload the app package
$app = Add-PnPApp -Path .\sharepoint\solution\pzl-quick-links.sppkg -Scope Site -Publish
# Install the web parts on the site
Install-PnPApp -Identity $app.Id -Scope Site -Wait
```

### Pages

* Create a page for the *All links* web part.
* Add the *Your links* web page to a page, and set the web part properties to point to the *all links* page

## Building

### Building the code

```bash
git clone the repo
npm i
npm i -g gulp
gulp bundle
```

### Testing
You can test by using the workbench
`gulp serve`


```
https://<tenant>.sharepoint.com/sites/<site>/_layouts/15/workbench.aspx?debug=true&noredir=true&debugManifestsFile=https://localhost:4321/temp/manifests.js
```

### Building the code for production

```bash
gulp bundle --ship
gulp package-solution --ship
```
