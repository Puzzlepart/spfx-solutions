import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
    BaseListViewCommandSet,
    Command,
    IListViewCommandSetListViewUpdatedParameters,
    IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import { Web, List, RenderListDataOptions } from '@pnp/sp';
import { HttpClient, HttpClientResponse } from '@microsoft/sp-http';
import * as JSZip from 'jszip';

import * as strings from 'PdfExportCommandSetStrings';

export interface IPdfExportCommandSetProperties {

}

interface SharePointFile {
    serverRelativeUrl: string;
    pdfUrl: string;
    fileType: string;
}

const LOG_SOURCE: string = 'PdfExportCommandSet';

export default class PdfExportCommandSet extends BaseListViewCommandSet<IPdfExportCommandSetProperties> {

    @override
    public onInit(): Promise<void> {
        Log.info(LOG_SOURCE, 'Initialized PdfExportCommandSet');
        return Promise.resolve();
    }

    @override
    public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
        const exportCommand: Command = this.tryGetCommand('EXPORT');
        if (exportCommand) {
            exportCommand.visible = event.selectedRows.length > 0;
        }
        const saveCommand: Command = this.tryGetCommand('SAVE_AS');
        if (saveCommand) {
            saveCommand.visible = event.selectedRows.length > 0;
        }
    }

    @override
    public async onExecute(event: IListViewCommandSetExecuteEventParameters): Promise<void> {
        switch (event.itemId) {
            case 'EXPORT':
                //Dialog.alert(`${this.properties.sampleTextOne}`);
                let zipFile: JSZip = new JSZip();
                
                break;
            case 'SAVE_AS':
                let item = event.selectedRows[0];
                let itemIds = event.selectedRows.map(i => i.getValueByName("ID"));
                //let itemId = parseInt(item.getValueByName("ID"));
                let files = await this.generatePdfUrls(itemIds);
                await this.saveAsPdf(files);
                break;
            default:
                throw new Error('Unknown command');
        }
    }

    private async saveAsPdf(files: SharePointFile[]) {
        let web: Web = new Web(this.context.pageContext.web.absoluteUrl);
        for (let i = 0; i < files.length; i++) {
            const file = files[i];
            let pdfUrl = file.serverRelativeUrl.replace("." + file.fileType, ".pdf");
            await web.getFileByServerRelativeUrl(file.serverRelativeUrl).copyTo(pdfUrl);
            let response = await this.context.httpClient.get(file.pdfUrl, HttpClient.configurations.v1);
            let blob = await response.blob();
            await web.getFileByServerRelativeUrl(pdfUrl).setContentChunked(blob);
        }
        window.location.href = window.location.href;
    }

    private async generatePdfUrls(listItemIds: string[]): Promise<SharePointFile[]> {
        let web: Web = new Web(this.context.pageContext.web.absoluteUrl);
        let options: RenderListDataOptions = RenderListDataOptions.EnableMediaTAUrls | RenderListDataOptions.ContextInfo | RenderListDataOptions.ListData | RenderListDataOptions.ListSchema;

        var values = listItemIds.map(i => { return `<Value Type='Counter'>${i}</Value>`; });

        const viewXml: string = `
        <View Scope='RecursiveAll'>
            <Query>
                <Where>
                    <In>
                        <FieldRef Name='ID' />
                        <Values>
                            ${values.join("")}
                        </Values>
                    </In>
                </Where>
            </Query>
            <RowLimit>${listItemIds.length}</RowLimit>
        </View>`;


        let response = await web.lists.getById(this.context.pageContext.list.id.toString()).renderListDataAsStream({ RenderOptions: options, ViewXml: viewXml });
        // "{.mediaBaseUrl}/transform/pdf?provider=spo&inputFormat={.fileType}&cs={.callerStack}&docid={.spItemUrl}&{.driveAccessToken}"
        let pdfConversionUrl = response.ListSchema[".pdfConversionUrl"];
        let mediaBaseUrl = response.ListSchema[".mediaBaseUrl"];
        let callerStack = response.ListSchema[".callerStack"];
        let driveAccessToken = response.ListSchema[".driveAccessToken"];

        let pdfUrls: SharePointFile[] = [];
        response.ListData.Row.forEach(element => {
            let fileType = element[".fileType"]
            let spItemUrl = element[".spItemUrl"];
            let pdfUrl = pdfConversionUrl
                .replace("{.mediaBaseUrl}", mediaBaseUrl)
                .replace("{.fileType}", fileType)
                .replace("{.callerStack}", callerStack)
                .replace("{.spItemUrl}", spItemUrl)
                .replace("{.driveAccessToken}", driveAccessToken);
            //console.log(pdfUrl);
            pdfUrls.push({ serverRelativeUrl: element["FileRef"], pdfUrl: pdfUrl, fileType: fileType });
        });
        return pdfUrls;
    }
}
