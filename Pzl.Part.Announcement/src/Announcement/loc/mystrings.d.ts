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
  EditDialogTitle: string
  SaveButtonText: string
  CancelButtonText: string
  CloseButtonText: string
  RequiredFieldText: string
  CreateSuccessText: string
  EditSuccessText: string
  CreateErrorText: string
  ChoicePlaceholder: string
  PeoplePickerPlaceholder: string
  PeoplePickerSearchingText: string
  PeoplePickerNoResultsText: string
  Banner: {
    DismissTitle: string
    OpenDetails: string
  }
  Status: {
    Planned: string
    Ongoing: string
    Resolved: string
  }
  QuickActions: {
    MarkOngoing: string
    MarkResolved: string
    Edit: string
    StatusChangedText: string
    StatusChangeErrorText: string
  }
  Periods: {
    Active: string
    Upcoming: string
    Past: string
  }
  Table: {
    Title: string
    Status: string
    Level: string
    From: string
    To: string
  }
  Editor: {
    SectionTitle: string
  }
  Notification: {
    SectionTitle: string
    AlwaysNotifiedHint: string
    ChangeErrorText: string
  }
  DatePicker: {
    Placeholder: string
    OpenTitle: string
    TimeLabel: string
    HourLabel: string
    MinuteLabel: string
    OkButtonText: string
    ClearButtonText: string
    GoToToday: string
    PrevMonth: string
    NextMonth: string
    PrevYear: string
    NextYear: string
  }
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
  /**
   * Display names of the template fields, used when the list schema cannot be read
   */
  FieldTitles: Record<string, string>
}

declare module 'AnnouncementStrings' {
  const strings: IAnnouncementStrings
  export = strings
}
