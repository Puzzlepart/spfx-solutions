import { BaseApplicationCustomizer } from '@microsoft/sp-application-base'
import { getPageRedirectUrl } from '../../shared/deepLink'

/**
 * The customizer has no properties. The custom action is registered with an
 * empty `ClientSideComponentProperties`.
 */
export type IAnnouncementDeepLinkApplicationCustomizerProperties = Record<string, never>

/**
 * Sends the user to the page a `?driftsmelding=page:<url>` deep link points to
 * (news notifications from the Teams notification service). The Announcement
 * web part does the same when it renders, but it only runs when the section
 * holding it is rendered, which on small screens may be never or late. An
 * application customizer runs as soon as the page's SPFx components load,
 * regardless of layout, so the redirect happens before the content is shown.
 *
 * `driftsmelding:<id>` deep links (focus an announcement) are left to the web
 * part, which owns the popover.
 */
export default class AnnouncementDeepLinkApplicationCustomizer extends BaseApplicationCustomizer<IAnnouncementDeepLinkApplicationCustomizerProperties> {
  public onInit(): Promise<void> {
    const url = getPageRedirectUrl()
    if (url) window.location.replace(url)
    return Promise.resolve()
  }
}
