## Move Everyone to visitors by Puzzlepart

Extension which check on public O365 groups that Everyone is in the visitors security group and not
in the members security group - ensuring they have read-only access and not contributor.

If a Group switches from Public to Private and back, this extension will when the site is accessed by a site owner
make sure permissions are fixed.

### Add to a site using the following PnP template
```
<?xml version="1.0"?>
<pnp:Provisioning 
    xmlns:pnp="http://schemas.dev.office.com/PnP/2017/05/ProvisioningSchema">
    <pnp:Preferences Generator="OfficeDevPnP.Core, Version=2.19.1710.0, Culture=neutral, PublicKeyToken=null" />
    <pnp:Templates ID="CONTAINER-TEMPLATE-EVERYONE-VISITOR">
        <pnp:ProvisioningTemplate ID="TEMPLATE-EVERYONE-VISITOR" Version="1" BaseSiteTemplate="GROUP#0" Scope="RootSite">
            <pnp:CustomActions>
                <pnp:SiteCustomActions>
                    <pnp:CustomAction
                        Title="MoveEveryoneApplicationCustomizer"
                        Name="MoveEveryoneApplicationCustomizer"
                        Rights="ManageWeb"                        
                        Location="ClientSideExtension.ApplicationCustomizer"
                        ClientSideComponentId="0ec55642-6231-464b-a4c2-dc72cb61d6f4"
                        ClientSideComponentProperties="{}" />
                </pnp:SiteCustomActions>
            </pnp:CustomActions>
        </pnp:ProvisioningTemplate>
    </pnp:Templates>
</pnp:Provisioning>
```


### Building the code

```bash
git clone the repo
npm i
gulp --ship
gulp package-solution --ship
```

This package produces the following:

* sharepoint/solution/pzl-ext-moveeveryone.sppkg - package to install in the App Catalog