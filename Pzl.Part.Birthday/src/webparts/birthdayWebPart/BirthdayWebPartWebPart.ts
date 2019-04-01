import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import {sp} from '@pnp/sp';
import * as strings from 'BirthdayWebPartWebPartStrings';
import BirthdayWebPart from './components/BirthdayWebPart';
import { IBirthdayWebPartProps } from './components/IBirthdayWebPartProps';

export interface IBirthdayWebPartWebPartProps {
  description: string;
}

export default class BirthdayWebPartWebPart extends BaseClientSideWebPart<IBirthdayWebPartWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IBirthdayWebPartProps > = React.createElement(
      BirthdayWebPart,
      {
        description: this.properties.description
      }
    );

    ReactDom.render(element, this.domElement);
  }

  public onInit(): Promise<void> {

    return super.onInit().then(_ => {
      sp.setup({
        spfxContext: this.context,
      });
    });
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
