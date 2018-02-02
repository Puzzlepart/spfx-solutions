import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import { BaseApplicationCustomizer, PlaceholderContent, PlaceholderName } from '@microsoft/sp-application-base';
import { MSGraph, Functions } from '../services';
import TeamsButton from './TeamsButton';
import CreateTeamsDialog from './CreateTeamsDialog';
import './AppCustomizer.scss';

const LOG_SOURCE: string = 'EnableTeamsApplicationCustomizer';

export interface IEnableTeamsApplicationCustomizerProperties {
    autoCreate: boolean;
}


/** A Custom Action which can be run during execution of a Client Side Application */
export default class EnableTeamsApplicationCustomizer
    extends BaseApplicationCustomizer<IEnableTeamsApplicationCustomizerProperties> {

    private _topPlaceholder: PlaceholderContent | undefined;

    @override
    public onInit(): Promise<void> {
        // Extra check for siteadmin to ensure it's run by a Group owner
        let isSiteAdmin = this.context.pageContext.legacyPageContext.isSiteAdmin;
        if (isSiteAdmin) {
            let autoCreate = this.properties.autoCreate.toString().toLowerCase() === 'true';
            this.DoWork(autoCreate);
        }
        return Promise.resolve();
    }

    private async DoWork(autoCreate: boolean) {
        let hasTeam = false;
        let groupId = this.context.pageContext.legacyPageContext.groupId;

        let endPointInfo = await MSGraph.Get(this.context.graphHttpClient, `beta/groups/${groupId}/endpoints`);
        if (endPointInfo && endPointInfo.value) {
            let info = endPointInfo.value.find(element => { return element.providerName === 'Microsoft Teams'; });
            hasTeam = info != null;
        }

        const dialog: CreateTeamsDialog = new CreateTeamsDialog();
        try {
            if (!hasTeam && autoCreate) {
                dialog.message = "Please wait while we set up Microsoft Teams and add a navigation link...";
                dialog.show();
                await Functions.CreateTeam(this.context.graphHttpClient, groupId, this.context.pageContext.site.absoluteUrl);
                hasTeam = true;
            }
        } catch (error) {
            dialog.close();
            Log.error(LOG_SOURCE, error);
        }

        if (hasTeam) {
            await Functions.RemoveCustomizer(this.context.pageContext.site.absoluteUrl, this.componentId);
        } else {
            if (!this._topPlaceholder) {
                this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top, {});
                if (!this._topPlaceholder || !this._topPlaceholder.domElement) {
                    Log.info(LOG_SOURCE, 'The expected placeholder (Top) was not found.');
                    return;
                }
                Log.info(LOG_SOURCE, 'The expected placeholder (Top) was found. Rendering <NavigationContainer />');

                let buttonProps = { graphClient: this.context.graphHttpClient, groupId: groupId, siteUrl: this.context.pageContext.site.absoluteUrl, componentId: this.componentId };
                ReactDOM.render(React.createElement(TeamsButton, buttonProps, null), this._topPlaceholder.domElement);
            }
        }
    }
}
