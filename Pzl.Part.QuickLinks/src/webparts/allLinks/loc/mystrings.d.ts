declare interface IAllLinksWebPartStrings {
  PropertyPane: {
    MandatoryLinksTitleLabel: string
    RecommendedLinksTitleLabel: string
    YourLinksTitleLabel: string
    DefaultIcon: string
    YourLinksOnTop: string
    ListingByCategory: string
    CategoryTitleFieldLabel: string
  }
  ActionRemoveMandatory: string
  AddLabel: string
  CancelLabel: string
  MandatoryLinksLabel: string
  NewLinkLabel: string
  NoCategoryLabel: string
  RecommendedLinksLabel: string
  SaveErrorLabel: string
  SaveOkLabel: string
  SaveYourLinksLabel: string
  TitleLabel: string
  UrlValidationLabel: string
  YourLinksLabel: string
}

declare module 'AllLinksWebPartStrings' {
  const strings: IAllLinksWebPartStrings
  export = strings
}
