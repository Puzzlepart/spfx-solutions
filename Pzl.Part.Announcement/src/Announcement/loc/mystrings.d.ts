declare interface IAnnouncementStrings {
  Aria: {
    HeaderInfoTitle: string
    PinnedTitle: string
    OpenDetails: string
    CloseDialog: string
  }
  AnnouncementsListName: string
  NotifiedAtText: string
  NoAnnouncementsText: string
  AnnouncementFetchErrorText: string
  YesText: string
  NoText: string
  CreateButtonText: string
  CreateDialogTitle: string
  SaveButtonText: string
  CancelButtonText: string
  RequiredFieldText: string
  CreateSuccessText: string
  CreateErrorText: string
  ChoicePlaceholder: string
  PeoplePickerPlaceholder: string
  PeoplePickerSearchingText: string
  PeoplePickerNoResultsText: string
  RichText: {
    ToolbarLabel: string
    Bold: string
    Italic: string
    Underline: string
    BulletList: string
    NumberList: string
    Link: string
    LinkPrompt: string
  }
  Severity: {
    Info: string
    Warning: string
    Error: string
    Success: string
  }
  /**
   * Display names of the template fields, used when the list schema cannot be read
   */
  FieldTitles: Record<string, string>
}

declare module 'AnnouncementStrings' {
  const strings: IAnnouncementStrings
  export = strings
}
