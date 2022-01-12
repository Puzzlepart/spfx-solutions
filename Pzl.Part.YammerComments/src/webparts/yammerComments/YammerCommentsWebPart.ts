import * as React from 'react';
import * as ReactDom from 'react-dom';
import * as packageSolution from '../../../config/package-solution.json';
import * as strings from 'YammerCommentsWebPartStrings';
import { AadTokenProvider, AadHttpClient } from '@microsoft/sp-http';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneLink,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown
} from '@microsoft/sp-property-pane';
import { PropertyPaneWebPartInformation } from '@pnp/spfx-property-controls/lib/PropertyPaneWebPartInformation';
import { YammerComments, IYammerCommentsProps } from './components/YammerComments';
import { YammerService, IYammerService } from './services/YammerService';

export interface IYammerCommentsWebPartProps {
  documentationUrl: string;
  community: IPropertyPaneDropdownOption;
}

export default class YammerCommentsWebPart extends BaseClientSideWebPart<IYammerCommentsWebPartProps> {

  private yammerService: IYammerService;

  private yammerCommunities: IPropertyPaneDropdownOption[];

  public async onInit(): Promise<void> {

    const tokenProvider: AadTokenProvider = await this.context.aadTokenProviderFactory.getTokenProvider();
    const aadHttpClient: AadHttpClient = await this.context.aadHttpClientFactory.getClient("https://api.yammer.com");
    this.yammerService = new YammerService(tokenProvider, aadHttpClient);

    console.log(this.context);
  }

  public render(): void {

    const element: React.ReactElement<IYammerCommentsProps> = React.createElement(
      YammerComments,
      {
        propertyPane: this.context.propertyPane,
        yammerService: this.yammerService,
        community: this.properties.community
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

  protected async loadPropertyPaneResources(): Promise<void> {
    if (!this.yammerCommunities) {
      try {
        var communities = await this.yammerService.getCommunities();
        this.yammerCommunities = new Array<IPropertyPaneDropdownOption>();
        communities.map(group => {
          this.yammerCommunities.push({ key: group.id, text: group.full_name });
        });
      } catch (error) {
        console.error(error);
      }
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.WebPartDescription
          },
          groups: [
            {
              groupFields: [
                PropertyPaneLink('', {
                  text: strings.DocumentationLinkLabel,
                  href: this.properties.documentationUrl,
                  target: "_blank"
                }),
                PropertyPaneDropdown('community', {
                  label: strings.CommunityFieldLabel,
                  options: this.yammerCommunities /*,
                  selectedKey: this.properties.community.key */
                })
              ]
            },
            {
              groupName: strings.WebPartAbout,
              groupFields: [
                PropertyPaneWebPartInformation({
                  description: `${strings.Version}: ${(<any>packageSolution).solution.version}`,
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
