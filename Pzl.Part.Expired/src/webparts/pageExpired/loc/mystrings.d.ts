declare interface IPageExpiredWebPartStrings {
  PropertyPaneDescription: string;

  ExpirationMessage: string;
  ExpireAfterLabel: string;
  
  PageWasPublished: string;

  DaysAgo: string;
  AMonthAgo: string;
  MonthsAgo: string;
  AYearAgo: string;
  YearsAgo: string;

  Verify: string;
  Ignore: string;
}

declare module 'PageExpiredWebPartStrings' {
  const strings: IPageExpiredWebPartStrings;
  export = strings;
}
