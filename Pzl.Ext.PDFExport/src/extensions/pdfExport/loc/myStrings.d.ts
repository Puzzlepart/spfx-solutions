declare interface IPdfExportCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'PdfExportCommandSetStrings' {
  const strings: IPdfExportCommandSetStrings;
  export = strings;
}
