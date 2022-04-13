declare interface IYammerEmbedWebPartStrings {
  WebPartDescription: string;  
  WebPartAbout:string;
  DocumentationLabel: string;
  EmbedWidgetLabel: string;
  PromptLabel: string;
  DefaultCommunityLabel: string;
  DefaultCommunityPlaceholder: string;
  Version: string;
}

declare module 'YammerEmbedWebPartStrings' {
  const strings: IYammerEmbedWebPartStrings;
  export = strings;
}
