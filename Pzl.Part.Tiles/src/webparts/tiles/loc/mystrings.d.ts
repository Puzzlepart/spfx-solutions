declare interface ITilesWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  CountFieldLabel: string;
  ImageWidthFieldLabel: string;
  ImageHeightFieldLabel: string;
  TextPaddingFieldLabel: string;
  DescriptionFieldLabel: string;
  DescriptionFieldDescription: string;
  BackgroundImageFieldLabel: string;
  BackgroundImageFieldDescription: string;
  FallbackImageUrlLabel: string;
  FallbackImageUrlDescription: string;
  NewTabFieldLabel: string;
  NewTabFieldDescription: string;
  LinkFieldLabel: string;
  LinkFieldDescription: string;
  OrderByFieldLabel: string;
  OrderByFieldDescription: string;
}

declare module 'TilesWebPartStrings' {
  const strings: ITilesWebPartStrings;
  export = strings;
}