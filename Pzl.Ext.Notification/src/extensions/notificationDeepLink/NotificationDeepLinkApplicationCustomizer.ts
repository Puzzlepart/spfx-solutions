import { BaseApplicationCustomizer } from '@microsoft/sp-application-base'
import { getPageRedirectUrl } from './pageDeepLink'

/**
 * The customizer has no properties. The custom action is registered with an
 * empty `ClientSideComponentProperties`.
 */
export type INotificationDeepLinkApplicationCustomizerProperties = Record<string, never>

/**
 * Sends the user to the page a `?driftsmelding=page:<url>` deep link points to
 * (news notifications from the Teams notification service). The notification
 * opens the intranet home page in the Teams tab; an application customizer
 * runs as soon as the page's SPFx components load, regardless of layout, so
 * the redirect happens before the content is shown.
 *
 * `driftsmelding:<id>` deep links (open an announcement) are left to the
 * Driftsmeldinger web part, which owns the popover.
 */
export default class NotificationDeepLinkApplicationCustomizer extends BaseApplicationCustomizer<INotificationDeepLinkApplicationCustomizerProperties> {
  public onInit(): Promise<void> {
    const url = getPageRedirectUrl()
    if (url) window.location.replace(url)
    return Promise.resolve()
  }
}
