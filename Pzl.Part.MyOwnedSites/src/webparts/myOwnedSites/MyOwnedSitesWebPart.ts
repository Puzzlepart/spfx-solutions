import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'MyOwnedSitesWebPartStrings';
import MyOwnedSites from './components/MyOwnedSites';
import { IMyOwnedSitesProps } from './components/IMyOwnedSitesProps';

export interface IMyOwnedSitesWebPartProps {
  description: string;
}

export default class MyOwnedSitesWebPart extends BaseClientSideWebPart<IMyOwnedSitesWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IMyOwnedSitesProps> = React.createElement(
      MyOwnedSites,
      {
        spfxContext: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    await super.onInit();
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
