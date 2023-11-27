import '@pnp/polyfill-ie11'
import * as React from 'react'
import * as ReactDom from 'react-dom'
import * as strings from 'QuickLinksWebPartStrings'
import { sp } from '@pnp/sp'
import { Version } from '@microsoft/sp-core-library'
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base'
import { ThemeProvider, ThemeChangedEventArgs, IReadonlyTheme } from '@microsoft/sp-component-base'
import {
  IPropertyPaneConfiguration,
  PropertyPaneSlider,
  PropertyPaneCheckbox,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane'
import { IQuickLinksProps } from './components/IQuickLinksProps'
import QuickLinks from './components/QuickLinks'

export interface IQuickLinksWebPartProps {
  title: string
  numberOfItems: number
  allLinksUrl: string
  defaultOfficeFabricIcon: string
  groupByCategory: boolean
  maxLinkLength: number
  lineHeight: number
  iconOpacity: number
  linkClickWebHook: string
}

export default class QuickLinksWebPart extends BaseClientSideWebPart<IQuickLinksWebPartProps> {
  private _themeProvidor: ThemeProvider // NOTE DO NOT REMOVE; we need to keep the reference for it not to (potentially) be garbage collected
  private _theme: IReadonlyTheme

  public render(): void {
    const element: React.ReactElement<IQuickLinksProps> = React.createElement(QuickLinks, {
      theme: this._theme,
      title: this.properties.title,
      userId: this.context.pageContext.legacyPageContext.userId,
      numberOfLinks: this.properties.numberOfItems,
      allLinksUrl: this.properties.allLinksUrl,
      defaultIcon: this.properties.defaultOfficeFabricIcon,
      groupByCategory: this.properties.groupByCategory,
      maxLinkLength: this.properties.maxLinkLength,
      lineHeight: this.properties.lineHeight,
      iconOpacity: this.properties.iconOpacity,
      webServerRelativeUrl: this.context.pageContext.web.serverRelativeUrl,
      linkClickWebHook: this.properties.linkClickWebHook
    })

    ReactDom.render(element, this.domElement)
  }
  public async onInit(): Promise<void> {
    sp.setup({
      spfxContext: this.context
    })

    const themeProvider: ThemeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey)
    this._theme = themeProvider.tryGetTheme()
    themeProvider.themeChangedEvent.add(this, this._handleThemeChange)
    this._themeProvidor = themeProvider

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
                PropertyPaneSlider('lineHeight', {
                  label: strings.propertyPane_LineHeightLabel,
                  min: 15,
                  max: 50
                }),
                PropertyPaneSlider('iconOpacity', {
                  label: strings.propertyPane_IconOpacityLabel,
                  min: 0,
                  max: 100
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
                }),
                PropertyPaneTextField('linkClickWebHook', {
                  label: strings.propertyPane_LinkClickWebHookLabel
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
