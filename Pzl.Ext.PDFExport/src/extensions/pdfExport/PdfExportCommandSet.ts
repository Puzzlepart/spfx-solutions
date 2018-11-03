import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
    BaseListViewCommandSet,
    Command,
    IListViewCommandSetListViewUpdatedParameters,
    IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import "@pnp/polyfill-ie11";
import { Web, List, RenderListDataOptions } from '@pnp/sp';
import { HttpClient, HttpClientResponse } from '@microsoft/sp-http';
import * as JSZip from 'jszip';
import * as FileSaver from 'file-saver';
import WaitDialog from './WaitDialog';
import * as strings from 'PdfExportCommandSetStrings';

export interface IPdfExportCommandSetProperties {
}

interface SharePointFile {
    serverRelativeUrl: string;
    pdfUrl: string;
    fileType: string;
    pdfFileName: string;
}

const LOG_SOURCE: string = 'PdfExportCommandSet';
const DIALOG = new WaitDialog({});

export default class PdfExportCommandSet extends BaseListViewCommandSet<IPdfExportCommandSetProperties> {

    private _validExts : string[] = ['csv', 'doc', 'docx', 'odp', 'ods', 'odt', 'pot', 'potm', 'potx', 'pps', 'ppsx', 'ppsxm', 'ppt', 'pptm', 'pptx', 'rtf', 'xls', 'xlsx'];

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

    // private toBuffer(ab) {
    //     var buffer = new Buffer(ab.byteLength);
    //     var view = new Uint8Array(ab);
    //     for (var i = 0; i < buffer.length; ++i) {
    //         buffer[i] = view[i];
    //     }
    //     return buffer;
    // }

    @override
    public async onExecute(event: IListViewCommandSetExecuteEventParameters): Promise<void> {
        let itemIds = event.selectedRows.map(i => i.getValueByName("ID"));
        let fileExts = event.selectedRows.map(i => i.getValueByName("File_x0020_Type").toLocaleLowerCase());
        DIALOG.showClose = false;
        for (let i = 0; i < fileExts.length; i++) {
            const ext = fileExts[i];
            if(this._validExts.indexOf(ext) === -1) {
                DIALOG.title = "Supported file extensions";
                DIALOG.message = "The current file extensions are supported: " + this._validExts.join(", ") + ".";
                DIALOG.error = "";
                DIALOG.showClose = true;
                DIALOG.show();
                return;
            }            
        }

        switch (event.itemId) {
            case 'EXPORT': {
                DIALOG.title = "Save as PDF";
                DIALOG.message = "Generating files...";
                DIALOG.error = "";
                DIALOG.show();
                let files = await this.generatePdfUrls(itemIds);
                if (itemIds.length == 1) {
                    const file = files[0];                    
                    DIALOG.message = `Processing ${file.pdfFileName}...`;
                    DIALOG.render();
                    const response = await this.context.httpClient.get(file.pdfUrl, HttpClient.configurations.v1);
                    if (response.ok) {
                        const blob = await response.blob();
                        FileSaver.saveAs(blob, file.pdfFileName);
                    } else {
                        const error = await response.json();
                        let errorMessage = error.error.innererror ? error.error.innererror.code : error.error.message;
                        DIALOG.error = `Failed to process ${file.pdfFileName} - ${errorMessage}`;
                        DIALOG.render();
                    }
                } else {
                    const zip: JSZip = new JSZip();
                    for (let i = 0; i < files.length; i++) {
                        const file = files[i];
                        DIALOG.message = `Processing ${file.pdfFileName}...`;
                        DIALOG.render();
                        const response = await this.context.httpClient.get(file.pdfUrl, HttpClient.configurations.v1);
                        if (response.ok) {
                            const blob = await response.blob();
                            zip.file(file.pdfFileName, blob, { binary: true });
                        } else {
                            const error = await response.json();
                            let errorMessage = error.error.innererror ? error.error.innererror.code : error.error.message;
                            DIALOG.error = `Failed to process ${file.pdfFileName} - ${errorMessage}`;
                            DIALOG.render();
                        }
                    }
                    let d = new Date();
                    let dateString = d.getFullYear() + "-" + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2);

                    const zipBlob = await zip.generateAsync({ type: "blob" });
                    FileSaver.saveAs(zipBlob, `files-${dateString}.zip`);
                }
                DIALOG.close();
                break;
            }
            case 'SAVE_AS': {
                DIALOG.title = "Generating PDF's";
                DIALOG.message = "Please wait while files are being processed...";
                DIALOG.show();
                let files = await this.generatePdfUrls(itemIds);
                await this.saveAsPdf(files);
                DIALOG.close();
                window.location.href = window.location.href;
                break;
            }
            default:
                throw new Error('Unknown command');
        }
    }

    private async saveAsPdf(files: SharePointFile[]) {
        let web: Web = new Web(this.context.pageContext.web.absoluteUrl);
        for (let i = 0; i < files.length; i++) {
            const file = files[i];
            DIALOG.message = `Processing ${file.pdfFileName}...`;
            DIALOG.render();
            let pdfUrl = file.serverRelativeUrl.replace("." + file.fileType, ".pdf");
            await web.getFileByServerRelativeUrl(file.serverRelativeUrl).copyTo(pdfUrl);
            let response = await this.context.httpClient.get(file.pdfUrl, HttpClient.configurations.v1);
            if (response.ok) {
                let blob = await response.blob();
                await web.getFileByServerRelativeUrl(pdfUrl).setContentChunked(blob);
            } else {
                const error = await response.json();
                let errorMessage = error.error.innererror ? error.error.innererror.code : error.error.message;
                DIALOG.error = `Failed to process ${file.pdfFileName} - ${errorMessage}`;
                DIALOG.render();
            }
        }
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
            let fileType = element[".fileType"];
            let spItemUrl = element[".spItemUrl"];
            let pdfUrl = pdfConversionUrl
                .replace("{.mediaBaseUrl}", mediaBaseUrl)
                .replace("{.fileType}", fileType)
                .replace("{.callerStack}", callerStack)
                .replace("{.spItemUrl}", spItemUrl)
                .replace("{.driveAccessToken}", driveAccessToken);
            let pdfFileName = element.FileLeafRef.replace(fileType, "pdf");
            pdfUrls.push({ serverRelativeUrl: element["FileRef"], pdfUrl: pdfUrl, fileType: fileType, pdfFileName: pdfFileName });
        });
        return pdfUrls;
    }
}
