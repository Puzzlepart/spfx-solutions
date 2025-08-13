import * as React from 'react'
import * as ReactDom from 'react-dom'
import * as strings from 'AllLinksWebPartStrings'
import { getSP } from '../pnpjsConfig'
import { Version } from '@microsoft/sp-core-library'
import { IAllLinksProps } from './components/types'
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base'
import { ThemeProvider, ThemeChangedEventArgs, IReadonlyTheme } from '@microsoft/sp-component-base'
import {
  IPropertyPaneConfiguration,
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
  hideYourLinks: boolean
  hideMandatoryLinks: boolean
  hideRecommendedLinks: boolean
}

export default class AllLinksWebPart extends BaseClientSideWebPart<IAllLinksWebPartProps> {
  private _themeProvider: ThemeProvider
  private _theme: IReadonlyTheme

  public render(): void {
    const element: React.ReactElement<IAllLinksProps> = React.createElement(AllLinks, {
      theme: this._theme,
      context: this.context,
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
      yourLinksDescription: this.properties.yourLinksDescription,
      hideYourLinks: this.properties.hideYourLinks,
      hideMandatoryLinks: this.properties.hideMandatoryLinks,
      hideRecommendedLinks: this.properties.hideRecommendedLinks
    } as IAllLinksProps)

    ReactDom.render(element, this.domElement)
  }

  public async onInit(): Promise<void> {
    getSP(this.context)

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
              groupName: strings.PropertyPane.ShowHideGroupName,
              isCollapsed: true,
              groupFields: [
                PropertyPaneToggle('hideYourLinks', {
                  label: strings.PropertyPane.HideYourLinksLabel
                }),
                PropertyPaneToggle('hideMandatoryLinks', {
                  label: strings.PropertyPane.HideMandatoryLinksLabel
                }),
                PropertyPaneToggle('hideRecommendedLinks', {
                  label: strings.PropertyPane.HideRecommendedLinksLabel
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
