import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import MyOwnedSites from './components/MyOwnedSites';
import { IMyOwnedSitesProps } from './components/IMyOwnedSitesProps';
import { IPropertyPaneConfiguration, PropertyPaneLabel, PropertyPaneToggle } from '@microsoft/sp-property-pane';
import { SPFI, spfi, SPFx } from "@pnp/sp";

export interface IMyOwnedSitesWebPartProps {
  includeSPSites: boolean;
}

export default class MyOwnedSitesWebPart extends BaseClientSideWebPart<IMyOwnedSitesWebPartProps> {
  private _sp: SPFI;

  public render(): void {
    const element: React.ReactElement<IMyOwnedSitesProps> = React.createElement(
      MyOwnedSites,
      {
        spfxContext: this.context,
        spClient: this._sp,
        includeSPSites: this.properties.includeSPSites
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    await super.onInit();
    this._sp = spfi().using(SPFx(this.context));
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
          groups: [
            {
              groupFields: [
                PropertyPaneToggle('includeSPSites', {
                  label: 'Include SharePoint sites',
                  checked: this.properties.includeSPSites
                }),
                PropertyPaneLabel('', {
                  text: `v${this.manifest.version}`
                })
              ]
            }
          ]
        }
      ]
    };
  }

}
