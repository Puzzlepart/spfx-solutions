declare interface IMyOwnedSitesWebPartStrings {
  SiteColumnName: string;
  DescriptionColumnName: string;
  CreatedDateColumnName: string;
  LoadingSpinnerLabel: string;
  GroupSitesTab: string;
  SitesTab: string;
}

declare module 'MyOwnedSitesWebPartStrings' {
  const strings: IMyOwnedSitesWebPartStrings;
  export = strings;
}
