declare interface IQuickLinksWebPartStrings {
  PropertyPane: {
    AllLinksUrlLabel: string
    DefaultOfficeFabricIconLabel: string
    DescriptionFieldDescription: string
    DescriptionFieldLabel: string
    GeneralGroupName: string
    GroupByCategoryLabel: string
    HideHeaderLabel: string
    HideShowAllLabel: string
    HideTitleLabel: string
    IconOpacityLabel: string
    LineHeightLabel: string
    LinkClickWebHookLabel: string
    MaxLinkLengthLabel: string
    NumberOfItemsLabel: string
    ShowHideGroupName: string
    TitleFieldDescription: string
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
