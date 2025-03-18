declare interface IAnnouncementStrings {
  Aria: {
    HeaderInfoTitle: string
  },
  AnnouncementsListName: string
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
