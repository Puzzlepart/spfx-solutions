import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IPropertyPaneConfiguration, PropertyPaneSlider, PropertyPaneCheckbox, PropertyPaneTextField } from "@microsoft/sp-property-pane";
import "@pnp/polyfill-ie11";
import { sp } from "@pnp/sp";
import * as strings from 'QuickLinksWebPartStrings';
import QuickLinks from './components/QuickLinks';
import { IQuickLinksProps } from './components/IQuickLinksProps';

export interface IQuickLinksWebPartProps {
  title: string;
  numberOfItems: number;
  allLinksUrl: string;
  defaultOfficeFabricIcon: string;
  groupByCategory: boolean;
  maxLinkLength: number;
}

export default class PzlQuickLinksWebPart extends BaseClientSideWebPart<IQuickLinksWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IQuickLinksProps> = React.createElement(
      QuickLinks,
      {
        title: this.properties.title,
        userId: this.context.pageContext.legacyPageContext.userId,
        numberOfLinks: this.properties.numberOfItems,
        allLinksUrl: this.properties.allLinksUrl,
        defaultIcon: this.properties.defaultOfficeFabricIcon,
        groupByCategory: this.properties.groupByCategory,
        maxLinkLength: this.properties.maxLinkLength,
        webServerRelativeUrl: this.context.pageContext.web.serverRelativeUrl
      }
    );

    ReactDom.render(element, this.domElement);
  }
  public async onInit(): Promise<void> {
    sp.setup({
      spfxContext: this.context,
    });
    try {
      await super.onInit();
      return;
    } catch (err) {
      return;
    }
  }
  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.propertyPane_TitleFieldLabel
                }),
                PropertyPaneSlider('numberOfItems', {
                  label: strings.propertyPane_NumberOfItemsLabel,
                  min: 0,
                  max: 500
                }),
                PropertyPaneSlider('maxLinkLength', {
                  label: strings.propertyPane_MaxLinkLengthLabel,
                  min: 50,
                  max: 500
                }),
                PropertyPaneTextField('allLinksUrl', {
                  label: strings.propertyPane_AllLinksUrlLabel
                }),
                PropertyPaneTextField('defaultOfficeFabricIcon', {
                  label: strings.propertyPane_DefaultOfficeFabricIconLabel
                }),
                PropertyPaneCheckbox('groupByCategory', {
                  text: strings.propertyPane_GroupByCategoryLabel,
                  checked: false
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
