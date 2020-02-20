declare interface ICustomTextEditorWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  ContentFieldLabel: string;
  StandardLabel: string;
  StandardLabelFade: string;
  AccordionLabel: string;
  BackgroundLabel: string;
  Show: string;
  LinkUnderline: string;
}

declare module 'CustomTextEditorWebPartStrings' {
  const strings: ICustomTextEditorWebPartStrings;
  export = strings;
}
