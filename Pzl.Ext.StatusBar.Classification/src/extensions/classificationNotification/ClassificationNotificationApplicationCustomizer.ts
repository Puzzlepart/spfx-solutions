import * as React from "react";
import * as ReactDOM from "react-dom";
import { override } from '@microsoft/decorators';
import {
    BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';
import ClassificationNotification, { IClassificationNotificationProps } from './ClassificationNotification/ClassificationNotification';
import * as strings from 'ClassificationNotificationApplicationCustomizerStrings';

const statusId: string = "PzlClassificationNotification";

export interface IClassificationNotificationApplicationCustomizerProperties {
    // This is an example; replace with your own property
    messageId: string;
    classifications: string;
}

export default class ClassificationNotificationApplicationCustomizer
    extends BaseApplicationCustomizer<IClassificationNotificationApplicationCustomizerProperties> {

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

        const element: React.ReactElement<IClassificationNotificationProps> = React.createElement(
            ClassificationNotification, { context: this.context, classifications: this.properties.classifications.split(',') }
        );

        ReactDOM.render(element, messageNode);
    }
}
