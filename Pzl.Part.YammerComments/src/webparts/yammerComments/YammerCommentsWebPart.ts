import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneLink,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { AadTokenProvider, AadHttpClient } from '@microsoft/sp-http';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { PropertyPaneWebPartInformation } from '@pnp/spfx-property-controls/lib/PropertyPaneWebPartInformation';
import * as strings from 'YammerCommentsWebPartStrings';
import YammerComments from './components/YammerComments';
import { IYammerCommentsProps } from './components/IYammerCommentsProps';
import * as packageSolution from '../../../config/package-solution.json';
import YammerService, { IYammerService } from './services/YammerService';


export interface IYammerCommentsWebPartProps {
  documentationUrl: string;
  communityId: string;
}

export default class YammerCommentsWebPart extends BaseClientSideWebPart<IYammerCommentsWebPartProps> {

  private yammerService: IYammerService;

  public async onInit(): Promise<void> {

      const tokenProvider: AadTokenProvider = await this.context.aadTokenProviderFactory.getTokenProvider();
      const aadHttpClient: AadHttpClient = await this.context.aadHttpClientFactory.getClient("https://api.yammer.com");
      this.yammerService = new YammerService(tokenProvider, aadHttpClient);
  }

  public render(): void {

    const element: React.ReactElement<IYammerCommentsProps> = React.createElement(
      YammerComments,
      {
        yammerService: this.yammerService,
        communityId: this.properties.communityId
      }
    );
    ReactDom.render(element, this.domElement);
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
              groupFields: [
                PropertyPaneTextField('communityId', {
                  label: strings.CommunityFieldLabel
                })
              ]
            },
            {
              groupName: strings.PropertyPaneGroupAbout,
              groupFields: [
                PropertyPaneWebPartInformation({
                  description: `${strings.Version}: ${(<any>packageSolution).solution.version}`,
                  key: 'version'
                }),
                PropertyPaneLink('', {
                  text: strings.DocumentationLinkLabel,
                  href: this.properties.documentationUrl,
                  target: "_blank"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
