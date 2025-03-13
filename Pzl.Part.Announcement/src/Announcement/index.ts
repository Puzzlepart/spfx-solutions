import * as React from 'react'
import * as ReactDom from 'react-dom'
import { Version } from '@microsoft/sp-core-library'
import { type IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane'
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base'
import { Announcement, IAnnouncementProps } from './components'

export default class AnnouncementWebPart extends BaseClientSideWebPart<IAnnouncementProps> {
  protected async onInit(): Promise<void> {
    await super.onInit()
  }

  public render(): void {
    const element: React.ReactElement<IAnnouncementProps> = React.createElement(Announcement, {
      ...this.properties,
      hasTeamsContext: !!this.context.sdks.microsoftTeams
    })

    ReactDom.render(element, this.domElement)
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
            }
          ]
        }
      ]
    }
  }
}
