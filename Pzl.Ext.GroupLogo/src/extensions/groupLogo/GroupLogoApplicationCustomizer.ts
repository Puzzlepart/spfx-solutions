import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
    BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import { SPHttpClient } from '@microsoft/sp-http';
import { Web, Site } from '@pnp/sp';
import { Dialog } from '@microsoft/sp-dialog';
import { MSGraph } from '../services';

import * as strings from 'GroupLogoApplicationCustomizerStrings';

export interface IGroupLogoApplicationCustomizerProperties {
    // This is an example; replace with your own property
    logoUrl: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class GroupLogoApplicationCustomizer
    extends BaseApplicationCustomizer<IGroupLogoApplicationCustomizerProperties> {

    @override
    public onInit(): Promise<void> {
        if (typeof console == "undefined" || typeof console.log == "undefined") var console = { log: function () { } };
        this.context.placeholderProvider.changedEvent.add(this, () => { this.DoWork(this.properties.logoUrl) });
        return Promise.resolve();
    }

    private async DoWork(logoUrl: string) {
        let isGroupOwner = this.context.pageContext.legacyPageContext.isSiteAdmin;
        if (!isGroupOwner) return;

        let response = await this.context.spHttpClient.post(`${logoUrl}/_api/contextinfo`, SPHttpClient.configurations.v1, {});
        if (!response.ok) return;
        let result = await response.json();
        // get web url from full path
        let webUrl = result.WebFullUrl;

        let web: Web = new Web(webUrl);
        let replace = `${window.location.protocol}//${window.location.hostname}`;
        let relativeUrlLogo = logoUrl.replace(replace, "");
        let buffer = await web.getFileByServerRelativeUrl(relativeUrlLogo).getBuffer();

        let hasError = false;
        let groupId = this.context.pageContext.legacyPageContext.groupId;

        try {
            Dialog.alert(strings.SettingUp);
            await MSGraph.Patch(this.context.graphHttpClient, `v1.0/groups/${groupId}/photo/$value`, buffer);
            console.log("Logo updated in the graph");
        } catch (err) {
            // Most likely due to user not having Exchange Online license or Group not ready
            console.log(err);
            hasError = true;
        }

        if (!hasError) {
            window.setTimeout(async () => {
                let currentWeb = new Web(this.context.pageContext.web.absoluteUrl);
                await currentWeb.getFolderByServerRelativeUrl(`${this.context.pageContext.web.serverRelativeUrl}/SiteAssets/__siteIcon__.jpg`).delete();
                console.log("Remove site icon file - to force update");
                this.removeCustomizer();
            }, 3000);
        }
    }

    private async removeCustomizer() {
        // Remove custom action from current sute
        let site = new Site(this.context.pageContext.site.absoluteUrl);
        let customActions = await site.userCustomActions.get();
        for (let i = 0; i < customActions.length; i++) {
            var instance = customActions[i];
            if (instance.ClientSideComponentId === this.componentId) {
                await site.userCustomActions.getById(instance.Id).delete();
                console.log("Logo extension removed");
                break;
            }
        }
    }
}
