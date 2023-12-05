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
  PropertyPaneLabel,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane'
import { AllLinks } from './components'
import { PropertyFieldIconPicker } from '@pnp/spfx-property-controls/lib/PropertyFieldIconPicker'

export interface IAllLinksWebPartProps {
  recommendedLinksTitle: string
  recommendedLinksDescription: string
  yourLinksTitle: string
  yourLinksDescription: string
  mandatoryLinksTitle: string
  mandatoryLinksDescription: string
  defaultIcon: string
  groupByCategory: boolean
}

export default class AllLinksWebPart extends BaseClientSideWebPart<IAllLinksWebPartProps> {
  private _themeProvider: ThemeProvider
  private _theme: IReadonlyTheme

  public render(): void {
    const element: React.ReactElement<IAllLinksProps> = React.createElement(AllLinks, {
      theme: this._theme,
      currentUserId: this.context.pageContext.legacyPageContext.userId,
      currentUserName: this.context.pageContext.user.displayName,
      defaultIcon: this.properties.defaultIcon,
      webServerRelativeUrl: this.context.pageContext.web.serverRelativeUrl,
      groupByCategory: this.properties.groupByCategory,
      mandatoryLinksTitle: this.properties.mandatoryLinksTitle,
      mandatoryLinksDescription: this.properties.mandatoryLinksDescription,
      recommendedLinksTitle: this.properties.recommendedLinksTitle,
      recommendedLinksDescription: this.properties.recommendedLinksDescription,
      yourLinksTitle: this.properties.yourLinksTitle,
      yourLinksDescription: this.properties.yourLinksDescription
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
            description: strings.PropertyPane.HeaderDescription
          },
          displayGroupsAsAccordion: true,
          groups: [
            {
              groupName: strings.PropertyPane.GeneralGroupName,
              groupFields: [
                PropertyPaneTextField('yourLinksTitle', {
                  label: strings.PropertyPane.YourLinksLabel,
                  placeholder: strings.PropertyPane.TitlePlaceholder
                }),
                PropertyPaneTextField('yourLinksDescription', {
                  placeholder: strings.PropertyPane.DescriptionPlaceholder
                }),
                PropertyPaneLabel('divider', {
                  text: ' '
                }),
                PropertyPaneTextField('mandatoryLinksTitle', {
                  label: strings.PropertyPane.MandatoryLinksLabel,
                  placeholder: strings.PropertyPane.TitlePlaceholder
                }),
                PropertyPaneTextField('mandatoryLinksDescription', {
                  placeholder: strings.PropertyPane.DescriptionPlaceholder
                }),
                PropertyPaneLabel('divider', {
                  text: ' '
                }),
                PropertyPaneTextField('recommendedLinksTitle', {
                  label: strings.PropertyPane.RecommendedLinksLabel,
                  placeholder: strings.PropertyPane.TitlePlaceholder
                }),
                PropertyPaneTextField('recommendedLinksDescription', {
                  placeholder: strings.PropertyPane.DescriptionPlaceholder
                }),
                PropertyPaneLabel('divider', {
                  text: ' '
                }),
                PropertyPaneToggle('groupByCategory', {
                  label: strings.PropertyPane.GroupByCategoryLabel
                })
              ]
            },
            {
              groupName: strings.PropertyPane.AdvancedGroupName,
              isCollapsed: true,
              groupFields: [
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
