declare interface IQuickLinksWebPartStrings {
  PropertyPane: {
    AdvancedGroupName: string
    AllLinksUrlLabel: string
    DefaultOfficeFabricIconLabel: string
    DescriptionFieldDescription: string
    DescriptionFieldLabel: string
    GeneralGroupName: string
    GroupByCategoryLabel: string
    HeaderDescription: string
    HideHeaderLabel: string
    HideShowAllLabel: string
    HideTitleLabel: string
    IconsOnlyLabel: string
    IconOpacityLabel: string
    LineHeightLabel: string
    LinkClickWebHookLabel: string
    NumberOfItemsLabel: string
    RenderShadowLabel: string
    ResponsiveButtonsLabel: string
    ShowHideGroupName: string
    StylingGroupName: string
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
