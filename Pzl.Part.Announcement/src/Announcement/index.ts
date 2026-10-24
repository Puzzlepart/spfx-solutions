import * as React from 'react'
import * as ReactDom from 'react-dom'
import { Version } from '@microsoft/sp-core-library'
import {
  type IPropertyPaneConfiguration,
  type IPropertyPaneField,
  PropertyPaneLabel,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane'
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base'
import { Announcement, IAnnouncementProps, IListField } from './components'
import { AnnouncementListService } from './services/AnnouncementListService'
import { DEFAULT_POPOVER_FIELDS } from './constants'

export default class AnnouncementWebPart extends BaseClientSideWebPart<IAnnouncementProps> {
  private _service: AnnouncementListService
  /**
   * List fields, loaded when the property pane opens
   */
  private _fields: IListField[]

  public render(): void {
    const element: React.ReactElement<IAnnouncementProps> = React.createElement(Announcement, {
      ...this.properties,
      context: this.context,
      service: this._service
    })

    ReactDom.render(element, this.domElement)
  }

  protected async onInit(): Promise<void> {
    await super.onInit()
    this._service = new AnnouncementListService(this.context)
    if (!this.properties.popoverFields) {
      this.properties.popoverFields = { ...DEFAULT_POPOVER_FIELDS }
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement)
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0')
  }

  protected onPropertyPaneConfigurationStart(): void {
    if (this._fields) return
    this._service
      .getFields()
      .then((fields) => {
        this._fields = fields
        this.context.propertyPane.refresh()
      })
      .catch(() => {
        // Keep the "loading" label
      })
  }

  /**
   * One toggle per list field for the popover
   */
  private _getPopoverFieldToggles(): IPropertyPaneField<any>[] {
    if (!this._fields) {
      return [PropertyPaneLabel('popoverFieldsLoading', { text: 'Laster felter fra lista …' })]
    }
    return this._fields
      .filter((field) => field.internalName !== 'Title')
      .map((field) =>
        PropertyPaneToggle(`popoverFields.${field.internalName}`, {
          label: field.title,
          checked: this.properties.popoverFields?.[field.internalName] === true,
          onText: 'Vis',
          offText: 'Skjul'
        })
      )
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
                    'Beskrivelse, dukker opp ved trykk på info-ikonet ved siden av tittel.',
                  multiline: true,
                  rows: 4
                }),
                PropertyPaneToggle('targetAudience', {
                  label: 'Aktiver målgruppeangivelse',
                  onText: 'På',
                  offText: 'Av'
                })
              ]
            },
            {
              groupName: 'Felter i detaljvisning',
              groupFields: [
                PropertyPaneLabel('popoverFieldsDescription', {
                  text: 'Velg hvilke felter fra lista som vises når man klikker på en driftsmelding. Tittel vises alltid.'
                }),
                ...this._getPopoverFieldToggles()
              ]
            },
            {
              groupName: 'Skjul/vis',
              groupFields: [
                PropertyPaneToggle('hideHeader', {
                  label: 'Skjul header',
                  onText: 'På',
                  offText: 'Av'
                }),
                PropertyPaneToggle('hideCreateButton', {
                  label: 'Skjul knapp for ny driftsmelding',
                  onText: 'På',
                  offText: 'Av'
                }),
                PropertyPaneLabel('hideCreateButtonDescription', {
                  text: 'Knappen vises bare for brukere som har rettighet til å legge til elementer i lista.'
                })
              ]
            }
          ]
        }
      ]
    }
  }
}
