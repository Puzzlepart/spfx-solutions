declare interface IGlobalNavigationApplicationCustomizerStrings {
  Title: string;
  DefaultNavToggleText: string;
  DefaultLoadErrorText: string;
  Field_AffectedSystems_Title: string;
  Field_Description_Title: string;
  Field_Consequence_Title: string;
  Field_Responsible_Title: string;
  Field_InfoLink_Title: string;
  Anchor_External_Title: string;
  Dialog_Contact_SecondaryText: string;
  Responsible_Hover_Title: string;
}

declare module 'GlobalNavigationApplicationCustomizerStrings' {
  const strings: IGlobalNavigationApplicationCustomizerStrings;
  export = strings;
}
