import * as React from 'react'
import * as ReactDom from 'react-dom'
import * as strings from 'QuickLinksWebPartStrings'
import { sp } from '@pnp/sp'
import { Version } from '@microsoft/sp-core-library'
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base'
import { ThemeProvider, ThemeChangedEventArgs, IReadonlyTheme } from '@microsoft/sp-component-base'
import {
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneSlider,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane'
import { IQuickLinksProps, QuickLinks } from './components'
import { PropertyFieldIconPicker } from '@pnp/spfx-property-controls/lib/PropertyFieldIconPicker'

export interface IQuickLinksWebPartProps {
  title: string
  description: string
  allLinksUrl: string
  defaultIcon: string
  groupByCategory: boolean
  lineHeight: number
  iconsOnly: boolean
  iconSize: number
  linkClickWebHook: string
  hideHeader: boolean
  hideTitle: boolean
  hideShowAll: boolean
  renderShadow: boolean
  responsiveButtons: boolean
  allLinksText: string
  buttonAppearance: 'secondary' | 'primary' | 'outline' | 'subtle' | 'transparent'
}

export default class QuickLinksWebPart extends BaseClientSideWebPart<IQuickLinksWebPartProps> {
  private _themeProvider: ThemeProvider
  private _theme: IReadonlyTheme

  public render(): void {
    const element: React.ReactElement<IQuickLinksProps> = React.createElement(QuickLinks, {
      theme: this._theme,
      title: this.properties.title,
      description: this.properties.description,
      userId: this.context.pageContext.legacyPageContext.userId,
      allLinksUrl: this.properties.allLinksUrl,
      defaultIcon: this.properties.defaultIcon,
      groupByCategory: this.properties.groupByCategory,
      lineHeight: this.properties.lineHeight,
      iconsOnly: this.properties.iconsOnly,
      iconSize: this.properties.iconSize,
      webServerRelativeUrl: this.context.pageContext.web.serverRelativeUrl,
      linkClickWebHook: this.properties.linkClickWebHook,
      hideHeader: this.properties.hideHeader,
      hideTitle: this.properties.hideTitle,
      hideShowAll: this.properties.hideShowAll,
      renderShadow: this.properties.renderShadow,
      responsiveButtons: this.properties.responsiveButtons,
      buttonAppearance: this.properties.buttonAppearance,
      allLinksText: this.properties.allLinksText
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
            description: strings.PropertyPane.HeaderDescription
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
                PropertyPaneTextField('allLinksText', {
                  label: strings.PropertyPane.AllLinksTextFieldLabel,
                  description: strings.PropertyPane.AllLinksTextFieldDescription
                }),
                PropertyPaneToggle('groupByCategory', {
                  label: strings.PropertyPane.GroupByCategoryLabel
                })
              ]
            },
            {
              groupName: strings.PropertyPane.StylingGroupName,
              isCollapsed: false,
              groupFields: [
                PropertyPaneDropdown('buttonAppearance', {
                  label: strings.PropertyPane.ButtonAppearanceLabel,
                  selectedKey: 'subtle',
                  options: [
                    { key: 'secondary', text: 'Sekundær' },
                    { key: 'primary', text: 'Primær' },
                    { key: 'outline', text: 'Kontur' },
                    { key: 'subtle', text: 'Subtil' },
                    { key: 'transparent', text: 'Transparent' }
                  ]
                }),
                PropertyPaneSlider('lineHeight', {
                  label: strings.PropertyPane.LineHeightLabel,
                  step: 2,
                  min: 16,
                  max: 64
                }),
                PropertyPaneToggle('iconsOnly', {
                  label: strings.PropertyPane.IconsOnlyLabel
                }),
                PropertyPaneSlider('iconSize', {
                  label: strings.PropertyPane.IconSizeLabel,
                  value: QuickLinks.defaultProps.iconSize,
                  step: 4,
                  min: 12,
                  max: 32
                }),
                PropertyPaneToggle('responsiveButtons', {
                  label: strings.PropertyPane.ResponsiveButtonsLabel
                }),
                PropertyPaneToggle('renderShadow', {
                  label: strings.PropertyPane.RenderShadowLabel
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
            },
            {
              groupName: strings.PropertyPane.AdvancedGroupName,
              isCollapsed: true,
              groupFields: [
                PropertyPaneTextField('allLinksUrl', {
                  label: strings.PropertyPane.AllLinksUrlLabel
                }),
                PropertyFieldIconPicker('defaultIcon', {
                  currentIcon: this.properties.defaultIcon,
                  key: 'defaultIconId',
                  onSave: (icon: string) => {
                    this.properties.defaultIcon = icon
                  },
                  buttonLabel: strings.PropertyPane.SelectDefaultIconLabel,
                  renderOption: 'panel',
                  properties: this.properties,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  label: strings.PropertyPane.DefaultIconLabel
                }),
                PropertyPaneTextField('linkClickWebHook', {
                  label: strings.PropertyPane.LinkClickWebHookLabel
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
