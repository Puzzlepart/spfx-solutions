## Enable Teams by Puzzlepart

[![Enable Teams - Application Customizer](https://img.youtube.com/vi/10ftqnidVrE/0.jpg)](https://www.youtube.com/watch?v=10ftqnidVrE)
<br/>_YouTube Video_

Extension which will create a Team for the group and add a navigation link to the Team.
If the extension is initialized with `autoCreate=true` then a Team is automatically created instead of showing a "Get Teams" link.

The extension when added with `Rights="ManageWeb"` will only run for Group owners, which is a requirement.

### Add to a site using the following PnP template

The template takes the following PnP input setting as a boolean value `AutoCreate`.

```
<?xml version="1.0"?>
<pnp:Provisioning 
    xmlns:pnp="http://schemas.dev.office.com/PnP/2017/05/ProvisioningSchema">
    <pnp:Preferences Generator="OfficeDevPnP.Core, Version=2.19.1710.0, Culture=neutral, PublicKeyToken=null" />
    <pnp:Templates ID="CONTAINER-TEMPLATE-ENABLE-TEAMS">
        <pnp:ProvisioningTemplate ID="TEMPLATE-ENABLE-TEAMS" Version="1" BaseSiteTemplate="GROUP#0" Scope="RootSite">
            <pnp:CustomActions>
                <pnp:SiteCustomActions>
                    <pnp:CustomAction
                        Title="EnableTeamsApplicationCustomizer"
                        Name="EnableTeamsApplicationCustomizer"
                        Rights="ManageWeb"
                        Location="ClientSideExtension.ApplicationCustomizer"
                        ClientSideComponentId="fa80f680-bda9-4c15-a8fe-4e86be0bf593"
                        ClientSideComponentProperties="{&quot;autoCreate&quot;:&quot;{parameter:AutoCreate}&quot;}" />
                </pnp:SiteCustomActions>
            </pnp:CustomActions>
        </pnp:ProvisioningTemplate>
    </pnp:Templates>
</pnp:Provisioning>
```


### Building the package

```bash
git clone the repo
npm i
gulp --ship
gulp package-solution --ship
```

This package produces the following:

* sharepoint/solution/pzl-ext-enable-teams.sppkg - package to install in the App Catalog
