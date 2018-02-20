## Office 365 Group Logo by Puzzlepart

Extension which will upload a new logo image for the Office 365 Group.
The extension takes a parameter `logoUrl` which should point to a URL for an image in JPEG format stored in SharePoint Online.

The extension when added with `Rights="ManageWeb"` will only run for Group owners, which is a requirement.

### Add to a site using the following PnP template

The template takes the following PnP input setting as a string value `LogoUrl`.

```
<?xml version="1.0"?>
<pnp:Provisioning 
    xmlns:pnp="http://schemas.dev.office.com/PnP/2017/05/ProvisioningSchema">
    <pnp:Preferences Generator="OfficeDevPnP.Core, Version=2.19.1710.0, Culture=neutral, PublicKeyToken=null" />
    <pnp:Templates ID="CONTAINER-TEMPLATE-GROUPS-LOGO">
        <pnp:ProvisioningTemplate ID="TEMPLATE-GROUPS-LOGO" Version="1" BaseSiteTemplate="GROUP#0" Scope="RootSite">
            <pnp:CustomActions>
                <pnp:SiteCustomActions>
                    <pnp:CustomAction
                        Title="GroupLogoApplicationCustomizer"
                        Name="GroupLogoApplicationCustomizer"
                        Rights="ManageWeb"
                        Location="ClientSideExtension.ApplicationCustomizer"
                        ClientSideComponentId="349f176f-ae79-4adf-a680-be7f187628d1"
                        ClientSideComponentProperties="{&quot;logoUrl&quot;:&quot;{parameter:LogoUrl}&quot;}" />
                </pnp:SiteCustomActions>
            </pnp:CustomActions>
        </pnp:ProvisioningTemplate>
    </pnp:Templates>
</pnp:Provisioning>
```

Sample command using PnP PowerShell
```
Apply-PnPProvisioningTemplate -Path template.xml -Parameters @{"LogoUrl"="https://contoso.sharepoint.com/SiteAssets/logo.jpg"}
```

### Building the package

```bash
git clone the repo
npm i
gulp --ship
gulp package-solution --ship
```

This package produces the following:

* sharepoint/solution/pzl-ext-group-logo.sppkg - package to install in the App Catalog
