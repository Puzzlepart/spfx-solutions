import * as React from 'react'
import * as ReactDom from 'react-dom'
import { Log } from '@microsoft/sp-core-library'
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base'
import strings from 'AnnouncementStrings'
import { AnnouncementListService } from '../../Announcement/services/AnnouncementListService'
import { fetchAnnouncements } from '../../Announcement/services/fetchAnnouncements'
import { DEFAULT_POPOVER_FIELDS } from '../../Announcement/constants'
import { IAnnouncement, IListField } from '../../Announcement/components/types'
import type { IAnnouncementBannerProps } from './components/AnnouncementBanner'

const LOG_SOURCE = 'AnnouncementBannerApplicationCustomizer'

/**
 * Properties of the custom action (`ClientSideComponentProperties`), see
 * `Templates/BannerCustomAction.xml`. All are optional.
 */
export interface IAnnouncementBannerApplicationCustomizerProperties {
  /**
   * Site that hosts the announcement list. Defaults to the current web, so the
   * same custom action works on the intranet site and (pointed at it) on the
   * other sites in the hub.
   */
  listSiteUrl?: string
  /**
   * Title of the announcement list. Defaults to "Driftsmeldinger".
   */
  listTitle?: string
  /**
   * Filter announcements on the list's audience targeting column
   */
  targetAudience?: boolean
  /**
   * Fields shown in the details dialog, keyed by internal name (`true` =
   * visible). Defaults to the web part's default popover fields.
   */
  popoverFields?: Record<string, boolean>
}

/**
 * Shows the active announcements marked "Vis som banner" (`PzlShowAsBanner`)
 * as message bars in the page's top placeholder. Each banner can be dismissed
 * (remembered in localStorage until the announcement changes) and clicked to
 * open a dialog with the same details as the web part's popover.
 *
 * The list is read first; the UI bundle is loaded only when there is
 * something to show, so pages without banner announcements pay for one list
 * query and nothing else.
 */
export default class AnnouncementBannerApplicationCustomizer extends BaseApplicationCustomizer<IAnnouncementBannerApplicationCustomizerProperties> {
  private _placeholder: PlaceholderContent | undefined
  private _props: IAnnouncementBannerProps | undefined

  public async onInit(): Promise<void> {
    try {
      const props = await this._load()
      if (!props) return
      this._props = props
      this.context.placeholderProvider.changedEvent.add(this, this._render)
      this._render()
    } catch (error) {
      Log.warn(LOG_SOURCE, `Banner not shown: ${error}`)
    }
  }

  /**
   * Reads the list and returns the props for the banner, or `null` when there
   * is no active announcement marked for the banner.
   */
  private async _load(): Promise<IAnnouncementBannerProps | null> {
    const { listSiteUrl, listTitle, targetAudience } = this.properties
    const service = new AnnouncementListService(
      this.context,
      listTitle?.trim() || strings.AnnouncementsListName,
      listSiteUrl?.trim() || undefined
    )
    const fields = await service.getFields()
    const announcements = (
      await fetchAnnouncements(service, this.context, fields, targetAudience === true)
    ).filter((announcement) => announcement.showAsBanner)
    if (!announcements.length) return null

    const popoverFields = this.properties.popoverFields ?? DEFAULT_POPOVER_FIELDS
    const detailFields = fields.filter((field) => popoverFields[field.internalName] === true)
    return { announcements, fields: detailFields }
  }

  private _render(): void {
    if (!this._props) return
    if (!this._placeholder) {
      this._placeholder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top, {
        onDispose: () => this._dispose()
      })
    }
    if (!this._placeholder) {
      Log.warn(LOG_SOURCE, 'The top placeholder is not available on this page')
      return
    }
    const element = this._placeholder.domElement
    import(/* webpackChunkName: 'announcement-banner-ui' */ './components/AnnouncementBanner')
      .then(({ AnnouncementBanner }) => {
        ReactDom.render(React.createElement(AnnouncementBanner, this._props), element)
      })
      .catch((error) => Log.warn(LOG_SOURCE, `Could not render the banner: ${error}`))
  }

  private _dispose(): void {
    if (this._placeholder) ReactDom.unmountComponentAtNode(this._placeholder.domElement)
    this._placeholder = undefined
  }

  protected onDispose(): void {
    this._dispose()
  }
}

export type { IAnnouncement, IListField }
