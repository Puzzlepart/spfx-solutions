declare interface IRssFeedWebPartStrings {
  PropertyPaneDescription: string;
  GeneralGroupName: string;
  HeaderTextFieldLabel: string;
  RssFeedUrlFieldLabel: string;
  ItemsCountFieldLabel: string;
  IconLabel: string;
  CacheExpirationTimeFieldLabel: string;
  Rss2jsonApiKeyFieldLabel: string;
  View_PublishLabel: string;
  View_EmptyPlaceholder_Label: string;
  View_EmptyPlaceholder_Description: string;
  View_EmptyPlaceholder_Button: string;
}

declare module 'RssFeedWebPartStrings' {
  const strings: IRssFeedWebPartStrings;
  export = strings;
}
