import * as React from 'react';
import * as ReactDOM from 'react-dom';

import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import { SPHttpClient, ISPHttpClientOptions } from '@microsoft/sp-http';
import {
    BaseListViewCommandSet,
    Command,
    IListViewCommandSetListViewUpdatedParameters,
    IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import PickDestination from './PickDestination';

import * as strings from 'MoveCopyPageCommandSetStrings';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IMoveCopyPageCommandSetProperties {
    // This is an example; replace with your own properties
    sampleTextOne: string;
    sampleTextTwo: string;
}

const LOG_SOURCE: string = 'MoveCopyPageCommandSet';

export default class MoveCopyPageCommandSet extends BaseListViewCommandSet<IMoveCopyPageCommandSetProperties> {

    @override
    public onInit(): Promise<void> {
        Log.info(LOG_SOURCE, 'Initialized MoveCopyPageCommandSet');
        return Promise.resolve();
    }

    @override
    public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
        const copyCommand: Command = this.tryGetCommand('COPYPAGE');
        if (copyCommand) {
            copyCommand.visible = event.selectedRows.length > 0;
        }
        const moveCommand: Command = this.tryGetCommand('MOVEPAGE');
        if (moveCommand) {
            moveCommand.visible = event.selectedRows.length > 0;
        }
    }

    @override
    public async onExecute(event: IListViewCommandSetExecuteEventParameters): Promise<void> {
        let itemIds = event.selectedRows.map(i => {
            let fileRelativePath = i.getValueByName("FileRef");
            return this.context.pageContext.site.absoluteUrl + fileRelativePath.replace(this.context.pageContext.site.serverRelativeUrl, "");
        });

        switch (event.itemId) {
            case 'COPYPAGE':
                //await this.fileOperation(itemIds);

                const element: React.ReactElement<any> = React.createElement(
                    PickDestination,
                    {
                    }
                );
                const container = document.createElement("div");
                document.body.appendChild(container);
                ReactDOM.render(element, container);
                //Dialog.alert(`${this.properties.sampleTextOne}`);
                break;
            case 'MOVEPAGE':
                Dialog.alert(`${this.properties.sampleTextTwo}`);
                break;
            default:
                throw new Error('Unknown command');
        }
    }

    public async fileOperation(fileUrls: string[]) {
        const currentSiteUrl = this.context.pageContext.web.absoluteUrl;
        const destinationUri = 'https://techmikael.sharepoint.com/teams/group-graphcall/SitePages';

        let bodyContent = {
            destinationUri,
            exportObjectUris: fileUrls,
            options: { AllowSchemaMismatch: true, IgnoreVersionHistory: true, IsMoveMode: false }
        };

        const request = await this.context.spHttpClient.post(`${currentSiteUrl}/_api/site/CreateCopyJobs`, SPHttpClient.configurations.v1, {
            body: JSON.stringify(bodyContent)
        });

        if (request.ok) {
            let result = await request.json();
            for (let i = 0; i < result.value.length; i++) {
                const jobInfo = result.value[i];
                this.checkStatus(jobInfo);
            }
        }
    }

    private async checkStatus(jobInfo) {
        const currentSiteUrl = this.context.pageContext.web.absoluteUrl;
        let checkBody = { "copyJobInfo": jobInfo };

        const opt: ISPHttpClientOptions = {
            body: JSON.stringify(checkBody)
        };

        const checkJob = await this.context.spHttpClient.post(`${currentSiteUrl}/_api/site/GetCopyJobProgress`,
            SPHttpClient.configurations.v1, opt);

        if (checkJob.ok) {
            let jobStatus = await checkJob.json();
            let jobState = jobStatus.JobState;
            if (jobState === 0) {
                let logs = jobStatus.Logs;
                for (let j = 0; j < logs.length; j++) {
                    const logEntry = JSON.parse(logs[j]);

                    if (logEntry.Event == "JobError") {
                        let message = logEntry.Message;
                    }
                    if (logEntry.Event == "JobFinishedObjectInfo") {
                        //done
                    }
                }
            } else {
                window.setTimeout(() => { this.checkStatus(jobInfo); }, 1000);
            }
        }
    }
}
