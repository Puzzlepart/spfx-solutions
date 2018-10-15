declare interface IPageNavigationWebPartStrings {

  // Property pane

  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  topLevelPageFieldLabel: string
  lookupFieldLabel: string;
  isRootExpanded: string;


}

declare module 'PageNavigationWebPartStrings' {
  const strings: IPageNavigationWebPartStrings;
  export = strings;
}
