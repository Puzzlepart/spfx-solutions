## External Access Disclaimer Status by Puzzlepart

![statusbar](./statusbar.png)

Extension which will show a message if
* The Group supports adding external members
* The modern site support sharing with external users

The extension takes a parameter `messageId` which should be a unique ID, and should match the ID used with `pzl-ext-status.sppkg`.

Requires:

* [Modern Site Status Renderer by Puzzlepart](../Pzl.Ext.StatusBar/README.md)

### Add to a site using the following PnP template

The [template](./template.xml) takes the following PnP input setting as a string value `messageId`.

```xml
<?xml version="1.0"?>
<pnp:Provisioning 
    xmlns:pnp="http://schemas.dev.office.com/PnP/2018/01/ProvisioningSchema">
    <pnp:Preferences Generator="OfficeDevPnP.Core, Version=2.19.1710.0, Culture=neutral, PublicKeyToken=null" />
    <pnp:Templates ID="CONTAINER-TEMPLATE-GROUPS-EXTERNAL">
        <pnp:ProvisioningTemplate ID="TEMPLATE-GROUPS-EXTERNAL" Version="1" BaseSiteTemplate="GROUP#0" Scope="RootSite">
            <pnp:CustomActions>
                <pnp:SiteCustomActions>
                    <pnp:CustomAction
                        Title="ExternalAccessDisclaimerApplicationCusto"
                        Name="ExternalAccessDisclaimerApplicationCusto"
                        Location="ClientSideExtension.ApplicationCustomizer"
                        ClientSideComponentId="2c565a15-92a5-4f51-9194-6a88a5edd482"
                        ClientSideComponentProperties="{&quot;messageId&quot;:&quot;{parameter:MessageId}&quot;}" />
                </pnp:SiteCustomActions>
            </pnp:CustomActions>
        </pnp:ProvisioningTemplate>
    </pnp:Templates>
</pnp:Provisioning>
```

Sample command using PnP PowerShell
```
Apply-PnPProvisioningTemplate -Path template.xml -Parameters @{"MessageId"="PzlMsg"}
```

### Building the package

```bash
git clone the repo
npm i
gulp --ship
gulp package-solution --ship
```

This package produces the following:

* sharepoint/solution/pzl-ext-external-disclaimer.sppkg - package to install in the App Catalog
