import * as React from 'react';
import { Log } from '@microsoft/sp-core-library';
import { IIconProps } from 'office-ui-fabric-react/lib/Icon';
import { Button, ButtonType } from 'office-ui-fabric-react/lib/Button';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { Functions } from '../services';
import { GraphHttpClient } from '@microsoft/sp-http';
import CreateTeamsDialog from './CreateTeamsDialog';

const LOG_SOURCE: string = 'EnableTeamsApplicationCustomizer';

export interface ITeamsButtonProps {
    graphClient: GraphHttpClient;
    groupId: string;
    siteUrl: string;
    componentId: string;
    shouldRedirect: boolean;
}

export default class TeamsButton extends React.Component<ITeamsButtonProps, { showPanel: boolean }> {

    constructor() {
        super();
        this.onClosePanel = this.onClosePanel.bind(this);
        this.onShowPanel = this.onShowPanel.bind(this);
        this.createTeam = this.createTeam.bind(this);
        this.state = {
            showPanel: false
        };
    }

    public render(): React.ReactElement<any> {
        let iconProps: IIconProps = { iconName: "TeamsLogo" };
        return (
            <div>
                <Button buttonType={ButtonType.primary} iconProps={iconProps} text="Get Teams!" className="teamsButton" onClick={this.onShowPanel} />
                <Panel
                    isOpen={this.state.showPanel}
                    type={PanelType.smallFixedFar}
                    onDismiss={this.onClosePanel}
                    headerText='Enable Microsoft Teams'
                    closeButtonAriaLabel='Close'
                >
                    <p>
                        Get ready for the Microsoft Teams experience!
                    </p>
                    <p>
                        Click the "Get Teams!" button below and you're ready to go in a few seconds.
                    </p>
                    <p>
                        <Button buttonType={ButtonType.compound}
                            className="panelButton" text="Get Teams!"
                            iconProps={iconProps}
                            onClick={this.createTeam}></Button>
                    </p>
                </Panel>
            </div>
        );
    }

    private onClosePanel(): void {
        this.setState({ showPanel: false });
    }

    private onShowPanel(): void {
        this.setState({ showPanel: true });
    }

    private async createTeam() {
        const dialog: CreateTeamsDialog = new CreateTeamsDialog();
        try {
            dialog.message = "Please wait while we set up Microsoft Teams and add a navigation link...";
            dialog.show();
            let teamsUri = await Functions.CreateTeam(this.props.graphClient, this.props.groupId, this.props.siteUrl);
            await Functions.RemoveCustomizer(this.props.siteUrl, this.props.componentId);

            if (this.props.shouldRedirect) {
                document.location.href = teamsUri;
            }
        } catch (error) {
            Log.error(LOG_SOURCE, error);
        } finally {
            dialog.close();
        }
    }
}