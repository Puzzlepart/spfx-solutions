declare interface IAnnouncementStrings {
  Aria: {
    HeaderInfoTitle: string
  }
  AnnouncementsListName: string
  AffectedSystemsLabel: string
  ConsequenceLabel: string
  ResponsibleLabel: string
  StartDateLabel: string
  EndDateLabel: string
  NoAnnouncementsText: string
  AnnouncementFetchErrorText: string
  Severity: {
    Info: string
    Warning: string
    Error: string
    Success: string
  }
}

declare module 'AnnouncementStrings' {
  const strings: IAnnouncementStrings
  export = strings
}
