declare interface IQuickLinksWebPartStrings {
  propertyPane_TitleFieldLabel: string;
  propertyPane_NumberOfItemsLabel: string;
  propertyPane_AllLinksUrlLabel: string;
  propertyPane_DefaultOfficeFabricIconLabel: string;
  propertyPane_GroupByCategoryLabel: string;
  propertyPane_MaxLinkLengthLabel: string;
  component_AllLinksLabel: string;
  component_NoCategoryLabel: string;
}

declare module 'QuickLinksWebPartStrings' {
  const strings: IQuickLinksWebPartStrings;
  export = strings;
}
