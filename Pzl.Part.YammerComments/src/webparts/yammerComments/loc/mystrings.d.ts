declare interface IYammerCommentsWebPartStrings {
  WebPartTitle: string;
  WebPartDescription: string;
  WebPartSetUp: string;
  WebPartAbout:string;
  DocumentationLinkLabel: string;
  CommunityFieldLabel: string;
  Version: string;
}

declare module 'YammerCommentsWebPartStrings' {
  const strings: IYammerCommentsWebPartStrings;
  export = strings;
}
