declare interface IQuickLinksWebPartStrings {
  PropertyPane: {
    AllLinksUrlLabel: string
    DefaultOfficeFabricIconLabel: string
    DescriptionFieldLabel: string
    GroupByCategoryLabel: string
    IconOpacityLabel: string
    LineHeightLabel: string
    LinkClickWebHookLabel: string
    MaxLinkLengthLabel: string
    NumberOfItemsLabel: string
    TitleFieldLabel: string
  }
  AllLinksLabel: string
  Description: string
  NoCategoryLabel: string
  Title: string
}

declare module 'QuickLinksWebPartStrings' {
  const strings: IQuickLinksWebPartStrings
  export = strings
}
