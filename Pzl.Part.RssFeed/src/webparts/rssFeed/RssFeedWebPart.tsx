import * as React from 'react';
import * as ReactDom from 'react-dom';
import * as packageSolution from '../../../config/package-solution.json';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IPropertyPaneConfiguration, PropertyPaneSlider, PropertyPaneTextField, PropertyPaneToggle } from "@microsoft/sp-property-pane";
import { PropertyPaneWebPartInformation } from '@pnp/spfx-property-controls/lib/PropertyPaneWebPartInformation';
import * as strings from 'RssFeedWebPartStrings';
import { RssFeed, IRssFeedProps } from './components/RssFeed';
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";

export interface IRssFeedWebPartProps {
  title: string;
  rssFeedUrl: string;
  itemsCount: number;
  officeUIFabricIcon: string;
  cacheDuration: number;
  apiKey: string;
  seeAllUrl: string;
  showItemDescription: boolean;
  showItemPubDate: boolean;
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
        showItemDescription: this.properties.showItemDescription,
        showItemPubDate: this.properties.showItemPubDate,
        updateProperty: (value: string) => {
          this.properties.title = value;
        },
        displayMode: this.displayMode,
        context: this.context,
        cacheDuration: this.properties.cacheDuration,
        instanceId: this.instanceId,
        seeAllUrl: this.properties.seeAllUrl,
      },
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
                PropertyPaneTextField('title', {
                  label: strings.WebPartTitle
                }),
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
                PropertyPaneToggle('showItemDescription', {
                  label: strings.ItemDescriptionLabel
                }),
                PropertyPaneToggle('showItemPubDate', {
                  label: strings.ItemPubDateLabel
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
            },
            {
              groupName: strings.WebPartAbout,
              groupFields: [
                PropertyPaneWebPartInformation({
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  description: `${strings.Version}: ${(packageSolution as any).solution.version}`,
                  key: 'version'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
