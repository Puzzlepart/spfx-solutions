declare interface IAllLinksWebPartStrings {
  PropertyPane: {
    MandatoryLinksTitleLabel: string
    RecommendedLinksTitleLabel: string
    YourLinksTitleLabel: string
    GroupByCategory: string
    CategoryTitleFieldLabel: string
    SelectDefaultIconLabel: string
    DefaultIconLabel: string
    GeneralGroupName: string
    HeaderDescription: string
    AdvancedGroupName: string
  }
  ActionRemoveMandatory: string
  AddLabel: string
  CancelLabel: string
  MandatoryLinksLabel: string
  MandatoryLinksDescription: string
  NewLinkLabel: string
  NoCategoryLabel: string
  RecommendedLinksLabel: string
  RecommendedLinksDescription: string
  SaveErrorLabel: string
  SaveOkLabel: string
  SaveYourLinksLabel: string
  TitleLabel: string
  TitlePlaceholder: string
  UrlPlaceholder: string
  UrlValidationLabel: string
  YourLinksLabel: string
  YourLinksDescription: string
}

declare module 'AllLinksWebPartStrings' {
  const strings: IAllLinksWebPartStrings
  export = strings
}
