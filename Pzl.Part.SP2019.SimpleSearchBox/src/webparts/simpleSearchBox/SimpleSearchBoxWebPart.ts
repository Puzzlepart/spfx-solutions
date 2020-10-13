import * as React from 'react';
import * as ReactDom from 'react-dom';
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneCheckbox } from "@microsoft/sp-webpart-base";
import * as strings from 'SimpleSearchBoxWebPartStrings';
import SimpleSearchBox from './components/SimpleSearchBox';
import { ISimpleSearchBoxProps } from './components/ISimpleSearchBoxProps';

export interface ISimpleSearchBoxWebPartProps {
  searchurl: string;
  title: string;
  openInNewTab: boolean;
  placeholder: string;
}

export default class SearchCentreWebPart extends BaseClientSideWebPart<ISimpleSearchBoxWebPartProps> {
  public render(): void {
    const element: React.ReactElement<ISimpleSearchBoxProps> = React.createElement(
      SimpleSearchBox,
      {
        searchurl: this.properties.searchurl,
        title: this.properties.title,
        openInNewTab: this.properties.openInNewTab,
        displayMode: this.displayMode,
        placeholder: this.properties.placeholder,
        updateProperty: (value: string) => {
          this.properties.title = value;
        }
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupFields: [
                PropertyPaneTextField('searchurl', {
                  label: strings.UrlFieldLabel
                }),
                PropertyPaneTextField('placeholder', {
                  label: strings.PlaceholderFieldLabel
                }),
                PropertyPaneCheckbox('openInNewTab', {
                  text: strings.OpenInNewTabLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
