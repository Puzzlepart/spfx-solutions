import * as packageSolution from '../../../config/package-solution.json';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneLink
} from '@microsoft/sp-property-pane';
import { PropertyPaneWebPartInformation } from '@pnp/spfx-property-controls/lib/PropertyPaneWebPartInformation';


import * as strings from 'YammerEmbedWebPartStrings';

export interface IYammerEmbedWebPartProps {
  documentationUrl: string;
  embedWidgetUrl: string;
  prompt: string;
  communityId: string;
}

export default class YammerEmbedWebPart extends BaseClientSideWebPart<IYammerEmbedWebPartProps> {

  private _isDarkTheme: boolean = false;

  protected onInit(): Promise<void> {
    return super.onInit();
  }

  public render(): void {

    const url = encodeURI(window.location.href);
    const prompt = encodeURIComponent(this.properties.prompt);
    const theme = this._isDarkTheme ? 'dark' : 'light';
    const defaultCommunity = this.properties.communityId ? `&defaultPublisherGroupId=${this.properties.communityId}` : '';

    console.log(defaultCommunity);

    const yammerEmbed = '<iframe name="embed-feed" title="Yammer" src="https://web.yammer.com/embed/attachable-link?header=false' +
      `&footer=false&theme=${theme}&promptText=${prompt}&url=${url}${defaultCommunity}` +
      '&showAttachableLinkPreview=false" style="border: 0px; overflow: hidden; width: 100%; height: 100%; min-height: 500px; resize: both;"></iframe>';

    this.domElement.innerHTML = yammerEmbed;
  }


  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;
    this.domElement.style.setProperty('--bodyText', semanticColors.bodyText);
    this.domElement.style.setProperty('--link', semanticColors.link);
    this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered);

  }

  protected onDispose(): void {

  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
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
                  text: strings.DocumentationLabel,
                  href: this.properties.documentationUrl,
                  target: "_blank"
                }),
                PropertyPaneTextField('prompt', {
                  label: strings.PromptLabel
                }),
                PropertyPaneTextField('communityId', {
                  label: strings.DefaultCommunityLabel,
                  placeholder: strings.DefaultCommunityPlaceholder
                }),
                PropertyPaneLink('', {
                  text: strings.EmbedWidgetLabel,
                  href: this.properties.embedWidgetUrl,
                  target: "_blank"
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
