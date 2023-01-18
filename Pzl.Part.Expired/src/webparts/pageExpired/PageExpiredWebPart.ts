import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneToggle } from '@microsoft/sp-property-pane';
import { SPPermission } from "@microsoft/sp-page-context";
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'PageExpiredWebPartStrings';
import { PageExpired, IPageExpiredProps } from './components/PageExpired';
import { IPageService, PageService } from './services/PageService';

export interface IPageExpiredWebPartProps {
  expireAfter: number;
  showForReaders: boolean;
}

export default class PageExpiredWebPart extends BaseClientSideWebPart<IPageExpiredWebPartProps> {

  private _pageService: IPageService;

  private _modified: Date;
  private _isEditor: boolean;
  private _isNews: boolean;

  public onVerify = async (event: unknown): Promise<void> => {
    await this._pageService.savePage();
    window.location.reload();
  }

  public render(): void {
    const element: React.ReactElement<IPageExpiredProps> = React.createElement(
      PageExpired,
      {
        verify: this.onVerify,
        modified: this._modified,
        expireAfter: Number(this.properties.expireAfter),
        showForReaders: this.properties.showForReaders,
        isEditor: this._isEditor,
        isNews: this._isNews
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    await super.onInit();
    this._pageService = new PageService(this.context.serviceScope);
    const page = await this._pageService.getPage();
    this._modified = new Date(page.Modified);
    this._isEditor = (new SPPermission(this.context.pageContext.web.permissions.value)).hasPermission(SPPermission.addListItems)
    this._isNews = page.PromotedState === 2;
    return Promise.resolve();
  }


  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {

    if (!currentTheme) {
      return;
    }

    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

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
                PropertyPaneTextField('expireAfter', {
                  label: strings.ExpireAfterLabel
                }),
                PropertyPaneToggle('showForReaders', {
                  label: strings.MessageAudienceLabel,
                  offText: strings.EditorsOnly,
                  onText: strings.EditorsAndReaders
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
