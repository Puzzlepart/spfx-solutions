import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IPropertyPaneConfiguration, PropertyPaneCheckbox, PropertyPaneTextField } from "@microsoft/sp-property-pane";
import * as strings from 'AllLinksWebPartStrings';
import AllLinks from './components/AllLinks';
import { IAllLinksProps } from './components/IAllLinksProps';
import "@pnp/polyfill-ie11";
import { sp } from "@pnp/sp";

export interface IAllLinksWebPartProps {
  defaultOfficeFabricIcon: string;
  mylinksOnTop: boolean;
  listingByCategory: boolean;
  listingByCategoryTitle: string;
}

export default class AllLinksWebPart extends BaseClientSideWebPart<IAllLinksWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IAllLinksProps> = React.createElement(
      AllLinks,
      {
        currentUserId: this.context.pageContext.legacyPageContext.userId,
        currentUserName: this.context.pageContext.user.displayName,
        defaultIcon: this.properties.defaultOfficeFabricIcon,
        webServerRelativeUrl: this.context.pageContext.web.serverRelativeUrl,
        mylinksOnTop: this.properties.mylinksOnTop,
        listingByCategory: this.properties.listingByCategory,
        listingByCategoryTitle: this.properties.listingByCategoryTitle
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
          header: {
            description: ""
          },
          groups: [
            {
              groupFields: [
                PropertyPaneTextField('defaultOfficeFabricIcon', {
                  label: strings.propertyPane_defaultOfficeFabricIcon
                }),
                PropertyPaneCheckbox('mylinksOnTop', {
                   text: strings.propertyPane_myLinksOnTop,
                   checked: false
                }),
                PropertyPaneCheckbox('listingByCategory', {
                  text: strings.propertyPane_listingByCategory,
                  checked: false
               }),
               PropertyPaneTextField('listingByCategoryTitle', {
                label: strings.propertyPane_CategoryTitleFieldLabel
              }),
              ]
            }
          ]
        }
      ]
    };
  }
}
