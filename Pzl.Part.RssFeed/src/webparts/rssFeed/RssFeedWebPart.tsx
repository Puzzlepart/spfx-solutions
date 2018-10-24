import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider
} from '@microsoft/sp-webpart-base';

import * as strings from 'RssFeedWebPartStrings';
import RssFeed from './components/RssFeed';
import { IRssFeedProps } from './components/IRssFeedProps';
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";

export interface IRssFeedWebPartProps {
  title: string;
  rssFeedUrl: string;
  itemsCount: number;
  officeUIFabricIcon: string;
  cacheDuration: number;
  apiKey: string;
  seeAllUrl: string;
}

export default class RssFeedWebPart extends BaseClientSideWebPart<IRssFeedWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IRssFeedProps> = React.createElement(
      RssFeed,
      {
        title: this.properties.title,
        rssFeedUrl: this.properties.rssFeedUrl,
        apiKey: this.properties.apiKey,
        itemsCount: this.properties.itemsCount,
        officeUIFabricIcon: this.properties.officeUIFabricIcon,
        updateProperty: (value: string) => {
          this.properties.title = value;
        },
        displayMode: this.displayMode,
        context: this.context,
        cacheDuration: this.properties.cacheDuration,
        instanceId: this.instanceId,
        seeAllUrl: this.properties.seeAllUrl,
      }
    );

    ReactDom.render((this.properties.rssFeedUrl) ? element : <Placeholder
      iconName='Edit'
      iconText={strings.View_EmptyPlaceholder_Label}
      description={strings.View_EmptyPlaceholder_Description}
      buttonLabel={strings.View_EmptyPlaceholder_Button}
      onConfigure={this._onConfigure.bind(this)} />, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private _onConfigure(): void {
    this.context.propertyPane.open();
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
              groupName: strings.GeneralGroupName,
              groupFields: [
                PropertyPaneTextField('seeAllUrl', {
                  label: strings.SeeAllUrlFieldLabel
                }),
                PropertyPaneTextField('rssFeedUrl', {
                  label: strings.RssFeedUrlFieldLabel
                }),
                PropertyPaneTextField('apiKey', {
                  label: strings.Rss2jsonApiKeyFieldLabel
                }),
                PropertyPaneTextField('officeUIFabricIcon', {
                  label: strings.IconLabel
                }),
                PropertyPaneSlider('itemsCount', {
                  label: strings.ItemsCountFieldLabel,
                  min: 1,
                  max: 20,
                }),
                PropertyPaneSlider('cacheDuration', {
                  label: strings.CacheExpirationTimeFieldDescription,
                  min: 0,
                  max: 1440
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}

