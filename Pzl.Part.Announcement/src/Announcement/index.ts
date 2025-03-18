import * as React from 'react'
import * as ReactDom from 'react-dom'
import { Version } from '@microsoft/sp-core-library'
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane'
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base'
import { Announcement, IAnnouncementProps } from './components'

export default class AnnouncementWebPart extends BaseClientSideWebPart<IAnnouncementProps> {
  public render(): void {
    const element: React.ReactElement<IAnnouncementProps> = React.createElement(Announcement, {
      ...this.properties,
      context: this.context
    })

    ReactDom.render(element, this.domElement)
  }

  protected async onInit(): Promise<void> {
    await super.onInit()
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement)
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0')
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: 'Konfigurasjon av webdelen.'
          },
          displayGroupsAsAccordion: true,
          groups: [
            {
              groupName: 'Generelt',
              groupFields: [
                PropertyPaneTextField('title', {
                  label: 'Tittel',
                  description: 'Tittel som vises i headeren på webdelen.'
                }),
                PropertyPaneTextField('description', {
                  label: 'Beskrivelse',
                  description:
                    'Beskrivelse av webdelen, dukker opp ved trykk på info-ikonet ved siden av tittel.',
                  multiline: true,
                  rows: 4
                })
              ]
            },
            {
              groupName: 'Skjul/vis',
              isCollapsed: true,
              groupFields: [
                PropertyPaneToggle('hideHeader', {
                  label: 'Skjul header',
                  onText: 'På',
                  offText: 'Av'
                })
              ]
            }
          ]
        }
      ]
    }
  }
}
