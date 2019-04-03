import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, DisplayMode } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import {sp} from '@pnp/sp';
import * as strings from 'BirthdayWebPartWebPartStrings';
import BirthdayWebPart from './components/BirthdayWebPart';
import { IBirthdayWebPartProps } from './components/IBirthdayWebPartProps';

export interface IBirthdayWebPartWebPartProps {
  title: string;
  itemsCount: number;
  displayMode: DisplayMode;
  context: WebPartContext;
  updateProperty: (value: string) => void;
}

export default class BirthdayWebPartWebPart extends BaseClientSideWebPart<IBirthdayWebPartWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IBirthdayWebPartProps > = React.createElement(
      BirthdayWebPart,
      {
        title: this.properties.title,
        itemsCount: this.properties.itemsCount,
        displayMode: this.displayMode,
        context: this.context,
        updateProperty: (value: string) =>{
          this.properties.title = value;
        }
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
                PropertyPaneTextField('title', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneSlider('itemsCount', {
                  label: strings.ItemsCountFieldLabel,
                  min: 1,
                  max: 20,
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
