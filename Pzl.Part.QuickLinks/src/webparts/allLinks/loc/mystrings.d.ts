declare interface IAllLinksWebPartStrings {
  PropertyPane: {
    AdvancedGroupName: string
    CategoryTitleFieldLabel: string
    DefaultIconLabel: string
    DescriptionPlaceholder: string
    GeneralGroupName: string
    GroupByCategoryLabel: string
    HeaderDescription: string
    MandatoryLinksLabel: string
    RecommendedLinksLabel: string
    SelectDefaultIconLabel: string
    TitlePlaceholder: string
    YourLinksLabel: string
  }
  ActionRemoveMandatory: string
  AddLabel: string
  CancelLabel: string
  IconLabel: string
  IconButtonLabel: string
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
