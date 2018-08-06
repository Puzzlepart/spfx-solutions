declare interface ITilesWebPartStrings {
  Property_PropertyPaneDescription: string;
  Property_BasicGroupName: string;
  Property_Count_Label: string;
  Property_ImageWidth_Label: string;
  Property_ImageHeight_Label: string;
  Property_TextPadding_Label: string;
  Property_Description_Label: string;
  Property_Description_Description: string;
  Property_BackgroundImage_Label: string;
  Property_BackgroundImage_Description: string;
  Property_FallbackImageUrl_Label: string;
  Property_FallbackImageUrl_Description: string;
  Property_NewTab_Label: string;
  Property_NewTab_Description: string;
  Property_Link_Label: string;
  Property_Link_Description: string;
  Property_OrderBy_Label: string;
  Property_OrderBy_Description: string;
  Property_TileType_Label: string;
  Property_TileChoiceField_Label: string;
  Property_List_Label: string;
  Property_AdvancedSettings_Label: string;
  Property_AdvancedSettings_OffText: string;
  Property_AdvancedSettings_OnText: string;
  View_NoOptionValue: string;
  View_EmptyPlaceholder_Label: string;
  View_EmptyPlaceholder_Description: string;
  View_EmptyPlaceholder_Button: string;
}

declare module 'TilesWebPartStrings' {
  const strings: ITilesWebPartStrings;
  export = strings;
}