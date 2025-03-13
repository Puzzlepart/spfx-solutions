declare interface IAnnouncementWebPartStrings {
  PropertyPaneDescription: string
  BasicGroupName: string
  DescriptionFieldLabel: string
  AppLocalEnvironmentSharePoint: string
  AppLocalEnvironmentTeams: string
  AppLocalEnvironmentOffice: string
  AppLocalEnvironmentOutlook: string
  AppSharePointEnvironment: string
  AppTeamsTabEnvironment: string
  AppOfficeEnvironment: string
  AppOutlookEnvironment: string
  UnknownEnvironment: string
}

declare module 'AnnouncementWebPartStrings' {
  const strings: IAnnouncementWebPartStrings;
  export = strings
}
