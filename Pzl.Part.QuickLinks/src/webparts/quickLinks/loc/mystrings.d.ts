declare interface IQuickLinksWebPartStrings {
  propertyPane_LinkClickWebHookLabel: string;
  propertyPane_LineHeightLabel: string;
  propertyPane_TitleFieldLabel: string;
  propertyPane_NumberOfItemsLabel: string;
  propertyPane_AllLinksUrlLabel: string;
  propertyPane_DefaultOfficeFabricIconLabel: string;
  propertyPane_GroupByCategoryLabel: string;
  propertyPane_MaxLinkLengthLabel: string;
  propertyPane_IconOpacityLabel: string;
  component_AllLinksLabel: string;
  component_NoCategoryLabel: string;
}

declare module 'QuickLinksWebPartStrings' {
  const strings: IQuickLinksWebPartStrings;
  export = strings;
}
