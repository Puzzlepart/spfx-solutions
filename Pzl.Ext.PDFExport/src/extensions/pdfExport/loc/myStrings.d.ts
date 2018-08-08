declare interface IPdfExportCommandSetStrings {
    DownloadAsPdf: string;
    SaveAsPdf: string;
}

declare module 'PdfExportCommandSetStrings' {
    const strings: IPdfExportCommandSetStrings;
    export = strings;
}
