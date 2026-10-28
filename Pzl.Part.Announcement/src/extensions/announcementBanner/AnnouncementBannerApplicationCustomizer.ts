import * as React from 'react'
import * as ReactDom from 'react-dom'
import { Log } from '@microsoft/sp-core-library'
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base'
import { getPageLocation, IBannerTargeting, isTargeted } from './targeting'
import type { IBannerListOptions } from './loadBanner'
import type { IAnnouncementBannerProps } from './components/AnnouncementBanner'

const LOG_SOURCE = 'AnnouncementBannerApplicationCustomizer'

/**
 * Marker on `window` set by the first instance that shows the banner on a
 * page. A site can be registered both tenant-wide and with its own custom
 * action; the second instance then does nothing instead of rendering twice.
 */
const INSTANCE_MARKER = '__pzlAnnouncementBanner'

/**
 * Properties of the custom action (`ClientSideComponentProperties`), see
 * `Templates/BannerCustomAction.xml` and the Tenant Wide Extensions row (README).
 * All are optional, but with no targeting (`siteUrls`, `hubSiteUrls` or
 * `everywhere`) the banner is shown nowhere.
 */
export interface IAnnouncementBannerApplicationCustomizerProperties
  extends IBannerTargeting, IBannerListOptions {
  /**
   * Site that hosts the announcement list. Defaults to the current web, so a
   * site-level registration on the intranet needs nothing; a tenant-wide
   * registration points this at the intranet site.
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
 * The entry bundle only decides whether the page is targeted (`siteUrls`,
 * `hubSiteUrls`, `everywhere`), synchronously from the page context. On a
 * targeted page the list code is loaded and the list read; the UI bundle is
 * loaded only when there is something to show. Pages outside the targeting
 * pay for this file alone, which is what makes a tenant-wide registration
 * affordable.
 */
export default class AnnouncementBannerApplicationCustomizer extends BaseApplicationCustomizer<IAnnouncementBannerApplicationCustomizerProperties> {
  private _placeholder: PlaceholderContent | undefined
  private _props: IAnnouncementBannerProps | undefined

  public async onInit(): Promise<void> {
    try {
      const page = getPageLocation(this.context.pageContext)
      if (!isTargeted(this.properties, page)) {
        Log.verbose(LOG_SOURCE, `Not targeted: ${page.webUrl}`)
        return
      }
      const w = window as any
      if (w[INSTANCE_MARKER]) {
        Log.info(LOG_SOURCE, 'Another registration already shows the banner on this page')
        return
      }
      w[INSTANCE_MARKER] = true

      const { loadBanner } = await import(
        /* webpackChunkName: 'announcement-banner-data' */ './loadBanner'
      )
      const props = await loadBanner(this.context, this.properties)
      if (!props) return
      this._props = props
      this.context.placeholderProvider.changedEvent.add(this, this._render)
      this._render()
    } catch (error) {
      Log.warn(LOG_SOURCE, `Banner not shown: ${error}`)
    }
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
    const w = window as any
    if (w[INSTANCE_MARKER]) delete w[INSTANCE_MARKER]
  }
}
