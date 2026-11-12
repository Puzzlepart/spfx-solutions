import * as React from 'react'
import { createRoot, type Root } from 'react-dom/client'
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
import { AnnouncementListService, isFormField } from './services/AnnouncementListService'
import { DEFAULT_POPOVER_FIELDS, isFormFieldEnabled } from './constants'

/**
 * Delay (ms) before a changed «Område for driftsmeldinger» is used, so the
 * list is not read for every key stroke in the property pane
 */
const LIST_SITE_URL_DELAY = 800

/**
 * Normalized `listSiteUrl`: trimmed, no trailing slash, empty means the current site
 */
function normalizeSiteUrl(value: string | undefined): string {
  return (value ?? '').trim().replace(/\/+$/, '')
}

export default class AnnouncementWebPart extends BaseClientSideWebPart<IAnnouncementProps> {
  private _service: AnnouncementListService
  /**
   * The `listSiteUrl` the current service reads from. Also the React key, so
   * the component is mounted again (and the list read) when it changes.
   */
  private _serviceSiteUrl: string
  private _listSiteUrlTimer: number
  /**
   * List fields, loaded when the property pane opens
   */
  private _fields: IListField[]
  private _root: Root | undefined

  public render(): void {
    const element: React.ReactElement<IAnnouncementProps> = React.createElement(Announcement, {
      ...this.properties,
      key: this._serviceSiteUrl || 'current',
      context: this.context,
      service: this._service
    })

    if (!this._root) {
      this._root = createRoot(this.domElement)
    }
    this._root.render(element)
  }

  /**
   * Creates the list service for the site in `listSiteUrl` (the central list)
   */
  private _createService(): void {
    this._serviceSiteUrl = normalizeSiteUrl(this.properties.listSiteUrl)
    this._service = new AnnouncementListService(
      this.context,
      undefined,
      this._serviceSiteUrl || undefined
    )
    this._fields = undefined
  }

  protected async onInit(): Promise<void> {
    await super.onInit()
    this._createService()
    if (!this.properties.popoverFields) {
      this.properties.popoverFields = { ...DEFAULT_POPOVER_FIELDS }
    }
    if (!this.properties.formFields) {
      this.properties.formFields = {}
    }
  }

  protected onDispose(): void {
    window.clearTimeout(this._listSiteUrlTimer)
    this._root?.unmount()
    this._root = undefined
  }

  protected onPropertyPaneFieldChanged(propertyPath: string): void {
    if (propertyPath !== 'listSiteUrl') return
    window.clearTimeout(this._listSiteUrlTimer)
    this._listSiteUrlTimer = window.setTimeout(() => {
      if (normalizeSiteUrl(this.properties.listSiteUrl) === this._serviceSiteUrl) return
      this._createService()
      this.render()
      this.onPropertyPaneConfigurationStart()
      this.context.propertyPane.refresh()
    }, LIST_SITE_URL_DELAY)
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

  /**
   * One toggle per editable list field for the create form
   */
  private _getFormFieldToggles(): IPropertyPaneField<any>[] {
    if (!this._fields) {
      return [PropertyPaneLabel('formFieldsLoading', { text: 'Laster felter fra lista …' })]
    }
    return this._fields.filter(isFormField).map((field) =>
      PropertyPaneToggle(`formFields.${field.internalName}`, {
        label: field.title,
        checked:
          field.required || isFormFieldEnabled(field.internalName, this.properties.formFields),
        disabled: field.required,
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
                PropertyPaneTextField('listSiteUrl', {
                  label: 'Område for driftsmeldinger',
                  description:
                    'Full adresse til området med lista Driftsmeldinger, for eksempel https://tenant.sharepoint.com. Tom betyr dette området. Brukerne må ha lesetilgang til lista.',
                  onGetErrorMessage: (value) =>
                    !value?.trim() || /^https:\/\/[^/]+/i.test(value.trim())
                      ? ''
                      : 'Skriv inn en full adresse som starter med https://'
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
              groupName: 'Felter i skjema for ny driftsmelding',
              groupFields: [
                PropertyPaneLabel('formFieldsDescription', {
                  text: 'Velg hvilke felter som vises i skjemaet for ny driftsmelding. Påkrevde felter og felter som er skjult i lista kan ikke endres her.'
                }),
                ...this._getFormFieldToggles()
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
