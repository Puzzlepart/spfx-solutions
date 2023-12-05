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
            description: strings.PropertyPane.HeaderDescription
          },
          displayGroupsAsAccordion: true,
          groups: [
            {
              groupName: strings.PropertyPane.GeneralGroupName,
              groupFields: [
                PropertyPaneCheckbox('groupByCategory', {
                  text: strings.PropertyPane.GroupByCategory,
                  checked: false
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
            },
            {
              groupName: strings.PropertyPane.AdvancedGroupName,
              isCollapsed: true,
              groupFields: [
                // PropertyFieldIconPicker('defaultIcon', {
                //   currentIcon: this.properties.defaultIcon,
                //   key: 'defaultIconId',
                //   onSave: (icon: string) => {
                //     this.properties.defaultIcon = icon
                //   },
                //   buttonLabel: strings.PropertyPane.SelectDefaultIconLabel,
                //   renderOption: 'panel',
                //   properties: this.properties,
                //   onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                //   label: strings.PropertyPane.DefaultIconLabel
                // }),
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
