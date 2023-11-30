import * as React from 'react'
import * as ReactDom from 'react-dom'
import * as strings from 'AllLinksWebPartStrings'
import { sp } from '@pnp/sp'
import { Version } from '@microsoft/sp-core-library'
import { IAllLinksProps } from './components/types'
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base'
import { ThemeProvider, ThemeChangedEventArgs, IReadonlyTheme } from '@microsoft/sp-component-base'
import {
  IPropertyPaneConfiguration,
  PropertyPaneCheckbox,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane'
import { AllLinks } from './components'

export interface IAllLinksWebPartProps {
  recommendedLinksTitle: string
  yourLinksTitle: string
  mandatoryLinksTitle: string
  defaultIcon: string
  yourLinksOnTop: boolean
  listingByCategory: boolean
  listingByCategoryTitle: string
}

export default class AllLinksWebPart extends BaseClientSideWebPart<IAllLinksWebPartProps> {
  private _themeProvider: ThemeProvider // NOTE keeping reference so that we are sure it is not going to be garbage collected
  private _theme: IReadonlyTheme

  public render(): void {
    const element: React.ReactElement<IAllLinksProps> = React.createElement(AllLinks, {
      theme: this._theme,
      currentUserId: this.context.pageContext.legacyPageContext.userId,
      currentUserName: this.context.pageContext.user.displayName,
      defaultIcon: this.properties.defaultIcon,
      webServerRelativeUrl: this.context.pageContext.web.serverRelativeUrl,
      yourLinksOnTop: this.properties.yourLinksOnTop,
      listingByCategory: this.properties.listingByCategory,
      listingByCategoryTitle: this.properties.listingByCategoryTitle,
      mandatoryLinksTitle: this.properties.mandatoryLinksTitle,
      recommendedLinksTitle: this.properties.recommendedLinksTitle,
      yourLinksTitle: this.properties.yourLinksTitle
    } as IAllLinksProps)

    ReactDom.render(element, this.domElement)
  }

  public async onInit(): Promise<void> {
    sp.setup({ spfxContext: this.context })

    const themeProvider: ThemeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey)
    this._theme = themeProvider.tryGetTheme()
    themeProvider.themeChangedEvent.add(this, this._handleThemeChange)
    this._themeProvider = themeProvider

    try {
      await super.onInit()
      return
    } catch (err) {
      return
    }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0')
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: ''
          },
          groups: [
            {
              groupFields: [
                PropertyPaneTextField('defaultIcon', {
                  label: strings.PropertyPane.DefaultIcon
                }),
                PropertyPaneCheckbox('yourLinksOnTop', {
                  text: strings.PropertyPane.YourLinksOnTop,
                  checked: false
                }),
                PropertyPaneCheckbox('listingByCategory', {
                  text: strings.PropertyPane.ListingByCategory,
                  checked: false
                }),
                PropertyPaneTextField('listingByCategoryTitle', {
                  label: strings.PropertyPane.CategoryTitleFieldLabel
                }),
                PropertyPaneTextField('mandatoryLinksTitle', {
                  label: strings.PropertyPane.MandatoryLinksTitleLabel
                }),
                PropertyPaneTextField('recommendedLinksTitle', {
                  label: strings.PropertyPane.RecommendedLinksTitleLabel
                }),
                PropertyPaneTextField('yourLinksTitle', {
                  label: strings.PropertyPane.YourLinksTitleLabel
                })
              ]
            }
          ]
        }
      ]
    }
  }

  private _handleThemeChange = (args: ThemeChangedEventArgs): void => {
    this._theme = args.theme
    this.render()
  }
}
