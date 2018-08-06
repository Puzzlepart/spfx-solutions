import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import { Web } from '@pnp/sp';

const LOG_SOURCE: string = 'HighLightActivationApplicationCustomizer';

export interface IHighLightActivationApplicationCustomizerProperties {
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class HighLightActivationApplicationCustomizer
    extends BaseApplicationCustomizer<IHighLightActivationApplicationCustomizerProperties> {

    @override
    public onInit(): Promise<void> {
        Log.info(LOG_SOURCE, `Initialized Highlight Home`);
        this.DoWork();
        return Promise.resolve();
    }

    private async DoWork() {
        let web = new Web(this.context.pageContext.web.absoluteUrl);
        let updateResult = await web.getList(`${this.context.pageContext.web.serverRelativeUrl}/SitePages`)
            .fields.getByInternalNameOrTitle("DocIcon").update({
                ClientSideComponentId: "9c89c914-ae2c-4d7d-8d72-de3b72fbbe9f",
            });
        Log.info(LOG_SOURCE, `Added field customizer to DocIcon: ${updateResult.data}`);
        await this.removeCustomizer();
    }

    private async removeCustomizer() {
        // Remove custom action from current sute
        let web = new Web(this.context.pageContext.web.absoluteUrl);
        let customActions = await web.userCustomActions.get();
        for (let i = 0; i < customActions.length; i++) {
            var instance = customActions[i];
            if (instance.ClientSideComponentId === this.componentId) {
                await web.userCustomActions.getById(instance.Id).delete();
                Log.info(LOG_SOURCE, "Extension removed");
                break;
            }
        }
    }
}
