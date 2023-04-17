import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import { BaseListViewCommandSet, Command, IListViewCommandSetListViewUpdatedParameters, IListViewCommandSetExecuteEventParameters } from '@microsoft/sp-listview-extensibility';
import "@pnp/polyfill-ie11";
import { Web, RenderListDataOptions } from '@pnp/sp/presets/all';
import { HttpClient } from '@microsoft/sp-http';
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
  private readonly _validExts: string[] = ['html', 'csv', 'doc', 'docx', 'odp', 'ods', 'odt', 'pot', 'potm', 'potx', 'pps', 'ppsx', 'ppsxm', 'ppt', 'pptm', 'pptx', 'rtf', 'xls', 'xlsx'];

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
    const itemIds = event.selectedRows.map(i => i.getValueByName("ID"));
    const fileExts = event.selectedRows.map(i => i.getValueByName("File_x0020_Type").toLocaleLowerCase());

    DIALOG.showClose = false;
    DIALOG.error = "";
    for (let i = 0; i < fileExts.length; i++) {
      const ext = fileExts[i];
      if (this._validExts.indexOf(ext) === -1) {
        DIALOG.title = strings.ExtSupport;
        DIALOG.message = strings.CurrentExtSupport + ": " + this._validExts.join(", ") + ".";
        DIALOG.showClose = true;
        DIALOG.show();
        return;
      }
    }

    switch (event.itemId) {
      case 'EXPORT': {
        DIALOG.title = strings.DownloadAsPdf;
        DIALOG.message = `${strings.GeneratingFiles}...`;
        DIALOG.show();
        const files = await this.generatePdfUrls(itemIds);
        let isOk = true;
        if (itemIds.length == 1) {
          const file = files[0];
          DIALOG.message = `${strings.Processing} ${file.pdfFileName}...`;
          DIALOG.render();
          const response = await this.context.httpClient.get(file.pdfUrl, HttpClient.configurations.v1);
          if (response.ok) {
            const blob = await response.blob();
            FileSaver.saveAs(blob, file.pdfFileName);
          } else {
            const error = await response.json();
            const errorMessage = error.error.innererror ? error.error.innererror.code : error.error.message;
            DIALOG.error = `${strings.FailedToProcess} ${file.pdfFileName} - ${errorMessage}<br/>`;
            DIALOG.render();
            isOk = false;
          }
        } else {
          const zip: JSZip = new JSZip();
          for (let i = 0; i < files.length; i++) {
            const file = files[i];
            DIALOG.message = `${strings.Processing} ${file.pdfFileName}...`;
            DIALOG.render();
            const response = await this.context.httpClient.get(file.pdfUrl, HttpClient.configurations.v1);
            if (response.ok) {
              const blob = await response.blob();
              zip.file(file.pdfFileName, blob, { binary: true });
            } else {
              const error = await response.json();
              const errorMessage = error.error.innererror ? error.error.innererror.code : error.error.message;
              DIALOG.error = `${strings.FailedToProcess} ${file.pdfFileName} - ${errorMessage}<br/>`;
              DIALOG.render();
              isOk = false;
            }
          }
          if (isOk) {
            zip.file("Powered by Puzzlepart.txt", "https://www.puzzlepart.com/");
            const d = new Date();
            const dateString = d.getFullYear() + "-" + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2) + '-' + ('0' + d.getHours()).slice(-2) + '-' + ('0' + d.getMinutes()).slice(-2) + '-' + ('0' + d.getSeconds()).slice(-2);

            const zipBlob = await zip.generateAsync({ type: "blob" });
            FileSaver.saveAs(zipBlob, `files-${dateString}.zip`);
          }
        }

        if (!isOk) {
          DIALOG.showClose = true;
          DIALOG.render();
        } else {
          DIALOG.close();
        }

        break;
      }
      case 'SAVE_AS': {
        DIALOG.title = strings.SaveAsPdf;
        DIALOG.message = `${strings.GeneratingFiles}...`;
        DIALOG.show();
        const files = await this.generatePdfUrls(itemIds);
        const ok = await this.saveAsPdf(files);
        if (ok) {
          DIALOG.close();
        } else {
          DIALOG.showClose = true;
          DIALOG.render();
        }
        break;
      }
      default:
        throw new Error('Unknown command');
    }
  }

  private async saveAsPdf(files: SharePointFile[]): Promise<boolean> {
    const web = Web(this.context.pageContext.web.absoluteUrl);
    let isOk = true;
    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      DIALOG.message = `${strings.Processing} ${file.pdfFileName}...`;
      DIALOG.render();
      const pdfUrl = file.serverRelativeUrl.replace("." + file.fileType, ".pdf");
      let exists = true;
      try {
        await web.getFileByServerRelativePath(pdfUrl).get();
        DIALOG.error += `${file.pdfFileName} ${strings.Exists}.<br/>`;
        DIALOG.render();
        isOk = false;
      } catch (error) {
        exists = false;
      }
      if (!exists) {
        const response = await this.context.httpClient.get(file.pdfUrl, HttpClient.configurations.v1);
        if (response.ok) {
          const blob = await response.blob();
          await web.getFileByServerRelativeUrl(file.serverRelativeUrl).copyTo(pdfUrl);
          await web.getFileByServerRelativeUrl(pdfUrl).setContentChunked(blob);
          const item = await web.getFileByServerRelativeUrl(pdfUrl).getItem("File_x0020_Type");
          // Potential fix for edge cases where file type is not set correctly
          if (item["File_x0020_Type"] !== "pdf") {
            await item.update({ "File_x0020_Type": "pdf" });
          }
        } else {
          const error = await response.json();
          const errorMessage = error.error.innererror ? error.error.innererror.code : error.error.message;
          DIALOG.error += `${strings.FailedToProcess}s ${file.pdfFileName} - ${errorMessage}<br/>`;
          DIALOG.render();
          isOk = false;
        }
      }
    }
    return isOk;
  }

  private async generatePdfUrls(listItemIds: string[]): Promise<SharePointFile[]> {
    const web = Web(this.context.pageContext.web.absoluteUrl);
    const options: RenderListDataOptions = RenderListDataOptions.EnableMediaTAUrls | RenderListDataOptions.ContextInfo | RenderListDataOptions.ListData | RenderListDataOptions.ListSchema;

    const values = listItemIds.map(i => { return `<Value Type='Counter'>${i}</Value>`; });

    const viewXml: string = (
      `<View Scope='RecursiveAll'>
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
        </View>`
    );

    const response = await web.lists.getById(this.context.pageContext.list.id.toString()).renderListDataAsStream({ RenderOptions: options, ViewXml: viewXml }) as any;
    // "{.mediaBaseUrl}/transform/pdf?provider=spo&inputFormat={.fileType}&cs={.callerStack}&docid={.spItemUrl}&{.driveAccessToken}"
    const pdfConversionUrl = response.ListSchema[".pdfConversionUrl"];
    const mediaBaseUrl = response.ListSchema[".mediaBaseUrl"];
    const callerStack = response.ListSchema[".callerStack"];
    const driveAccessToken = response.ListSchema[".driveAccessToken"];

    const pdfUrls: SharePointFile[] = [];
    response.ListData.Row.forEach((element) => {
      const fileType: string = element[".fileType"];
      const spItemUrl: string = element[".spItemUrl"];
      const pdfUrl: string = pdfConversionUrl
        .replace("{.mediaBaseUrl}", mediaBaseUrl)
        .replace("{.fileType}", fileType)
        .replace("{.callerStack}", callerStack)
        .replace("{.spItemUrl}", spItemUrl)
        .replace("{.driveAccessToken}", driveAccessToken);
      const pdfFileName: string = element.FileLeafRef.replace(fileType, "pdf");
      pdfUrls.push({ serverRelativeUrl: element["FileRef"], pdfUrl: pdfUrl, fileType: fileType, pdfFileName: pdfFileName });
    });
    return pdfUrls;
  }
}
