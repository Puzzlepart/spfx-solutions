declare interface IWelcomeMessageWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  WelcomeTextFieldLabel: string;
  WelcomeTextPlaceholder: string;
  RemoveWebPartMarginPaddingFieldLabel: string;
  On: string;
  Off: string;
}

declare module 'WelcomeMessageWebPartStrings' {
  const strings: IWelcomeMessageWebPartStrings;
  export = strings;
}
