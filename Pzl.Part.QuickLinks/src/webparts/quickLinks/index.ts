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
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane'
import { IQuickLinksProps, QuickLinks } from './components'

export interface IQuickLinksWebPartProps {
  title: string
  description: string
  numberOfItems: number
  allLinksUrl: string
  defaultOfficeFabricIcon: string
  groupByCategory: boolean
  maxLinkLength: number
  lineHeight: number
  iconOpacity: number
  linkClickWebHook: string
  hideHeader: boolean
  hideTitle: boolean
  hideShowAll: boolean
}

export default class QuickLinksWebPart extends BaseClientSideWebPart<IQuickLinksWebPartProps> {
  private _themeProvidor: ThemeProvider // NOTE DO NOT REMOVE; we need to keep the reference for it not to (potentially) be garbage collected
  private _theme: IReadonlyTheme

  public render(): void {
    const element: React.ReactElement<IQuickLinksProps> = React.createElement(QuickLinks, {
      theme: this._theme,
      title: this.properties.title,
      description: this.properties.description,
      userId: this.context.pageContext.legacyPageContext.userId,
      numberOfLinks: this.properties.numberOfItems,
      allLinksUrl: this.properties.allLinksUrl,
      defaultIcon: this.properties.defaultOfficeFabricIcon,
      groupByCategory: this.properties.groupByCategory,
      maxLinkLength: this.properties.maxLinkLength,
      lineHeight: this.properties.lineHeight,
      iconOpacity: this.properties.iconOpacity,
      webServerRelativeUrl: this.context.pageContext.web.serverRelativeUrl,
      linkClickWebHook: this.properties.linkClickWebHook,
      hideHeader: this.properties.hideHeader,
      hideTitle: this.properties.hideTitle,
      hideShowAll: this.properties.hideShowAll
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
          header: {
            description: 'Konfigurasjon av webdelen'
          },
          displayGroupsAsAccordion: true,
          groups: [
            {
              groupName: strings.PropertyPane.GeneralGroupName,
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.PropertyPane.TitleFieldLabel,
                  description: strings.PropertyPane.TitleFieldDescription
                }),
                PropertyPaneTextField('description', {
                  label: strings.PropertyPane.DescriptionFieldLabel,
                  description: strings.PropertyPane.DescriptionFieldDescription,
                  multiline: true
                }),
                PropertyPaneSlider('numberOfItems', {
                  label: strings.PropertyPane.NumberOfItemsLabel,
                  min: 0,
                  max: 500
                }),
                PropertyPaneSlider('maxLinkLength', {
                  label: strings.PropertyPane.MaxLinkLengthLabel,
                  min: 50,
                  max: 500
                }),
                PropertyPaneSlider('lineHeight', {
                  label: strings.PropertyPane.LineHeightLabel,
                  min: 15,
                  max: 50
                }),
                PropertyPaneSlider('iconOpacity', {
                  label: strings.PropertyPane.IconOpacityLabel,
                  min: 0,
                  max: 100
                }),
                PropertyPaneTextField('allLinksUrl', {
                  label: strings.PropertyPane.AllLinksUrlLabel
                }),
                PropertyPaneTextField('defaultOfficeFabricIcon', {
                  label: strings.PropertyPane.DefaultOfficeFabricIconLabel
                }),
                PropertyPaneCheckbox('groupByCategory', {
                  text: strings.PropertyPane.GroupByCategoryLabel,
                  checked: false
                }),
                PropertyPaneTextField('linkClickWebHook', {
                  label: strings.PropertyPane.LinkClickWebHookLabel
                })
              ]
            },
            {
              groupName: strings.PropertyPane.ShowHideGroupName,
              isCollapsed: true,
              groupFields: [
                PropertyPaneToggle('hideHeader', {
                  label: strings.PropertyPane.HideHeaderLabel
                }),
                !this.properties.hideHeader &&
                  PropertyPaneToggle('hideTitle', {
                    label: strings.PropertyPane.HideTitleLabel
                  }),
                !this.properties.hideHeader &&
                  PropertyPaneToggle('hideShowAll', {
                    label: strings.PropertyPane.HideShowAllLabel
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
