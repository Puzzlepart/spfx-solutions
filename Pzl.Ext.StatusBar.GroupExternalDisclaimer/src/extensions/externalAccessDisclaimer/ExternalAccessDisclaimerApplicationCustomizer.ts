import * as React from "react";
import * as ReactDOM from "react-dom";
import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
    BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import ExternalDisclaimer from './ExternalDisclaimer/ExternalDisclaimer';
import { IExternalDisclaimerProps } from './ExternalDisclaimer/ExternalDisclaimer';

import * as strings from 'ExternalAccessDisclaimerApplicationCustomizerStrings';
const LOG_SOURCE: string = 'ExternalAccessDisclaimerApplicationCustomizer';
const statusId: string = "PzlMsgExternalAccess";

export interface IExternalAccessDisclaimerApplicationCustomizerProperties {
    messageId: string;
}

export default class ExternalAccessDisclaimerApplicationCustomizer
    extends BaseApplicationCustomizer<IExternalAccessDisclaimerApplicationCustomizerProperties> {

    @override
    public onInit(): Promise<void> {
        this.context.placeholderProvider.changedEvent.add(this, this.render);
        return Promise.resolve();
    }

    private async render(): Promise<void> {

        let targetNode = document.getElementById(this.properties.messageId);
        if (!targetNode) {
            // Make sure message element is present
            targetNode = document.createElement("DIV");
            targetNode.id = this.properties.messageId;
            targetNode.style.cssText = "display:none";
            document.body.appendChild(targetNode);
        }

        let messageNode = document.getElementById(statusId);
        if (!messageNode) {
            // Make sure message element is present
            messageNode = document.createElement("DIV");
            messageNode.id = statusId;
            targetNode.appendChild(messageNode);
        }

        const element: React.ReactElement<IExternalDisclaimerProps> = React.createElement(
            ExternalDisclaimer, { context: this.context }
        );

        ReactDOM.render(element, messageNode);
    }
}
