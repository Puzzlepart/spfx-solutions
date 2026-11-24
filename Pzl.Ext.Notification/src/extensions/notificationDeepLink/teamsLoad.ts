import { QUERY_PARAM_NAME, getPageRedirectUrl } from './pageDeepLink'

/**
 * Teams keeps the tab's page alive (app caching) when the user opens the tab
 * again, for example by clicking another notification in the activity feed: the
 * page is not loaded again, so the customizer's check in `onInit` does not run.
 * Teams tells the page with a `load` message that carries the address the tab
 * would have been loaded with:
 *
 * `{ func: 'load', args: [{ entityId, contentUrl }] }`
 *
 * SharePoint answers the message but does not navigate, so the customizer
 * listens for it and sends the user to the page of the deep link. Kept free of
 * SPFx imports, like `pageDeepLink.ts`.
 */

/**
 * Host name endings of the applications that host Teams tabs
 */
const HOST_SUFFIXES = [
  'teams.microsoft.com',
  'teams.cloud.microsoft',
  'teams.microsoft.us',
  'teams.live.com',
  '.cloud.microsoft',
  '.office.com',
  '.microsoft365.com'
]

/**
 * Whether a message origin is an application that hosts Teams tabs
 *
 * @param origin Origin of the message event
 */
export function isTeamsHostOrigin(origin: string): boolean {
  try {
    const url = new URL(origin)
    if (url.protocol !== 'https:') return false
    return HOST_SUFFIXES.some((suffix) =>
      suffix.charAt(0) === '.'
        ? url.hostname.endsWith(suffix)
        : url.hostname === suffix || url.hostname.endsWith(`.${suffix}`)
    )
  } catch {
    return false
  }
}

/**
 * Path without a trailing slash, lower case, for comparing pages
 */
const normalizePath = (path: string): string => path.replace(/\/+$/, '').toLowerCase()

/**
 * Returns the URL to navigate to for a Teams `load` message, or `null` when
 * the message is something else or there is nowhere to go:
 *
 * - `page:<url>` (news): the news page, unless it is the current page.
 * - `driftsmelding:<id>` (announcement): the page of the deep link, with the
 *   parameter, when the user is on another page. On that page itself the
 *   Driftsmeldinger web part opens the announcement.
 *
 * Only addresses on the page's own origin are followed.
 *
 * @param data Data of the message event
 * @param location Location of the page (defaults to `window.location`)
 */
export function getTeamsLoadRedirectUrl(
  data: unknown,
  location: Pick<Location, 'origin' | 'pathname'> = window.location
): string | null {
  try {
    if (!data || typeof data !== 'object') return null
    const message = data as { func?: unknown; args?: unknown }
    if (message.func !== 'load' || !Array.isArray(message.args)) return null
    const contentUrl = (message.args[0] as { contentUrl?: unknown } | undefined)?.contentUrl
    if (typeof contentUrl !== 'string' || !contentUrl) return null

    // The tab loads teamslogon.aspx?dest=<page address>
    const tab = new URL(contentUrl)
    const dest = new URL(tab.searchParams.get('dest') || contentUrl)
    if (dest.origin !== location.origin) return null

    const page = getPageRedirectUrl({ pathname: location.pathname, search: dest.search })
    if (page) return page

    const value = dest.searchParams.get(QUERY_PARAM_NAME) || ''
    const isAnnouncement = /^(driftsmelding:)?\d+$/.test(value)
    if (isAnnouncement && normalizePath(dest.pathname) !== normalizePath(location.pathname)) {
      return `${dest.pathname}${dest.search}`
    }
    return null
  } catch {
    return null
  }
}

/**
 * Navigates when Teams shows the page again with a deep link. Returns a
 * function that stops listening.
 */
export function redirectOnTeamsLoad(): () => void {
  const listener = (event: MessageEvent): void => {
    if (!isTeamsHostOrigin(event.origin)) return
    const url = getTeamsLoadRedirectUrl(event.data)
    if (url) window.location.replace(url)
  }
  window.addEventListener('message', listener)
  return () => window.removeEventListener('message', listener)
}
