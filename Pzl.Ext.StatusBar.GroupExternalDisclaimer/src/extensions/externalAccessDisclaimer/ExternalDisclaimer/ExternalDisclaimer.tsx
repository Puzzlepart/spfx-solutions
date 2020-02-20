import * as React from 'react';
import { Icon } from "office-ui-fabric-react/lib/Icon";
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IWebPartContext } from '@microsoft/sp-webpart-base';
import { MSGraph } from '../services';
import { Log } from '@microsoft/sp-core-library';
import styles from './Styling.module.scss';
import * as strings from 'ExternalAccessDisclaimerApplicationCustomizerStrings';
import { Dialog } from '@microsoft/sp-dialog';

export interface IExternalDisclaimerProps {
    context: IWebPartContext;
}

export interface IExternalDisclaimerState {
    allowExternalMember: boolean;
    allowSharing: boolean;
    allowAnonymousSharing: boolean;
}
export default class ExternalDisclaimer extends React.PureComponent<IExternalDisclaimerProps, IExternalDisclaimerState> {
    public aadHttpClient;
    
    constructor(props) {
        super(props);
        this.state = {
            allowExternalMember: false,
            allowSharing: false,
            allowAnonymousSharing: false
        };
    }

    public async onInit(): Promise<void> {
        this.aadHttpClient = await this.context.aadHttpClientFactory.getClient('https://graph.microsoft.com');
    }

    public componentDidMount(): void {
        this.getExternalInfo();
    }

    public render(): React.ReactElement<null> {
        if (this.state.allowExternalMember || this.state.allowSharing || this.state.allowAnonymousSharing) {
            return (
                <div className={styles.warning}>
                    <Icon iconName='Warning' className={styles.warningIcon} /><span className={styles.warningAdjust}>{strings.ExternalAccess}</span>
                </div>
            );
        }
        return <span />;
    }

    private async getExternalInfo() {
        this.checkSharing().then((result) => {
            this.setState({
                allowSharing: result
            });
        });

        this.checkSharingAnonymous().then((result) => {
            this.setState({
                allowAnonymousSharing: result
            });
        });

        this.checkExternalMembers().then((result) => {
            this.setState({
                allowExternalMember: result
            });
        });
    }

    private async checkExternalMembers(): Promise<boolean> {
        try {
            let externalMembersTemplateId = "08d542b9-071f-4e16-94b0-74abb372e3d9";
            let groupId = this.props.context.pageContext.legacyPageContext.groupId;
            let graphUrl = `v1.0/groups/${groupId}/settings`;

            let groupSettings = await MSGraph.Get(this.aadHttpClient, graphUrl);
            let externalGuestsAllowed = false;
            for (var i = 0; i < groupSettings.value.length; i++) {
                let setting = groupSettings.value[i];
                if (setting.templateId === externalMembersTemplateId) {
                    for (var j = 0; j < setting.values.length; j++) {
                        var guestSetting = setting.values[j];
                        if (guestSetting.name === "AllowToAddGuests") {
                            externalGuestsAllowed = guestSetting.value.toLocaleLowerCase() === 'true';
                            break;
                        }
                    }
                }
            }
            return externalGuestsAllowed;

        } catch (error) {
            Log.error("Settings check error", error);
            return false;
        }
    }

    private async checkSharing(): Promise<boolean> {
        try {
            let url = `${this.props.context.pageContext.site.absoluteUrl}/_api/site/ShareByEmailEnabled`;

            let response: SPHttpClientResponse = await this.props.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
            let suggestions = await response.json();
            let allowed = suggestions.value;

            return allowed;

        } catch (error) {
            Log.error("Check sharing failed", error);
            return false;
        }
    }

    private async checkSharingAnonymous(): Promise<boolean> {
        try {
            let url = `${this.props.context.pageContext.site.absoluteUrl}/_api/site/ShareByLinkEnabled`;

            let response: SPHttpClientResponse = await this.props.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
            let suggestions = await response.json();
            let allowed = suggestions.value;
            return allowed;

        } catch (error) {
            Log.error("Check anonymous sharing failed", error);
            return false;
        }
    }
}
