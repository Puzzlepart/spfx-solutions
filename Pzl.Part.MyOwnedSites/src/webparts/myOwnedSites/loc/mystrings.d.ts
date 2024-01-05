declare interface IMyOwnedSitesWebPartStrings {
  SiteColumnName: string;
  DescriptionColumnName: string;
  CreatedDateColumnName: string;
  LoadingSpinnerLabel: string;
}

declare module 'MyOwnedSitesWebPartStrings' {
  const strings: IMyOwnedSitesWebPartStrings;
  export = strings;
}
