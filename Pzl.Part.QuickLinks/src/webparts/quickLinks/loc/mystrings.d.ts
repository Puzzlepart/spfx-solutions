declare interface IQuickLinksWebPartStrings {
  PropertyPane: {
    AdvancedGroupName: string
    AllLinksTextFieldDescription: string
    AllLinksTextFieldLabel: string
    AllLinksUrlLabel: string
    ButtonAppearanceLabel: string
    DefaultIconLabel: string
    DescriptionFieldDescription: string
    DescriptionFieldLabel: string
    GeneralGroupName: string
    GlobalConfigurationUrlLabel: string
    GlobalConfigurationUrlDescription: string
    GroupByCategoryLabel: string
    HeaderDescription: string
    HideHeaderLabel: string
    HideShowAllLabel: string
    HideTitleLabel: string
    IconSizeLabel: string
    IconsOnlyLabel: string
    LineHeightLabel: string
    GapSizeLabel: string
    LinkClickWebHookLabel: string
    RenderShadowLabel: string
    ResponsiveButtonsLabel: string
    SelectDefaultIconLabel: string
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
