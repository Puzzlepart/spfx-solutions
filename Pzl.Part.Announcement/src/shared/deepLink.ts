/**
 * Deep link parsing shared by the Announcement web part and the
 * AnnouncementDeepLink application customizer. Keep this module free of SPFx
 * imports: the customizer runs on every page load and should stay small.
 */

/**
 * Prefix used in the Teams deep link `subEntityId`/`subPageId`
 * (e.g. `driftsmelding:12`).
 */
export const SUB_PAGE_ID_PREFIX = 'driftsmelding:'

/**
 * Query string parameter used as fallback outside Teams
 * (e.g. `?driftsmelding=12`). The notification Teams app's tab also passes the
 * deep link `subEntityId` through this parameter, prefix included.
 */
export const QUERY_PARAM_NAME = 'driftsmelding'

/**
 * Prefix used in deep links for news pages (e.g. `page:/SitePages/News.aspx`).
 * The home page cannot show a news page, so the user is sent to it.
 */
export const PAGE_PREFIX = 'page:'

/**
 * Returns the server-relative page a `page:<url>` deep link points to, or
 * `null` when the parameter is absent or not a safe same-origin page path.
 * Only paths like `/sites/x/SitePages/Page.aspx` are accepted: no scheme,
 * no host, no `..`, no query or fragment.
 *
 * @param search Query string to read from (defaults to `window.location.search`)
 */
export function getPageDeepLink(search: string = window.location.search): string | null {
  try {
    const raw = new URLSearchParams(search).get(QUERY_PARAM_NAME)
    if (!raw || raw.indexOf(PAGE_PREFIX) !== 0) return null
    const path = raw.substring(PAGE_PREFIX.length)
    const lower = path.toLowerCase()
    const forbidden = ['?', '#', '..', String.fromCharCode(92)]
    if (path.charAt(0) !== '/' || path.charAt(1) === '/') return null
    if (forbidden.some((c) => path.indexOf(c) !== -1)) return null
    if (lower.substring(lower.length - 5) !== '.aspx') return null
    return path
  } catch {
    return null
  }
}

/**
 * Returns the URL to navigate to when the current page was opened with a
 * `page:<url>` deep link, or `null` when there is nothing to redirect to
 * (no deep link, or the deep link points to the current page). The
 * `driftsmelding` parameter is dropped and the other query parameters (e.g.
 * `app=portals` from the Teams tab) are carried over.
 *
 * @param location Location to read from (defaults to `window.location`)
 */
export function getPageRedirectUrl(
  location: Pick<Location, 'pathname' | 'search'> = window.location
): string | null {
  const page = getPageDeepLink(location.search)
  if (!page || page.toLowerCase() === location.pathname.toLowerCase()) return null
  const params = new URLSearchParams(location.search)
  params.delete(QUERY_PARAM_NAME)
  const query = params.toString()
  return query ? `${page}?${query}` : page
}

/**
 * Parses an announcement ID from a raw value. Accepts `driftsmelding:<id>`
 * or a bare numeric ID. Returns `null` for anything else (e.g. `page:...`
 * deep links for news).
 */
export function parseAnnouncementId(value: string | undefined | null): string | null {
  if (!value) return null
  const id =
    value.indexOf(SUB_PAGE_ID_PREFIX) === 0 ? value.substring(SUB_PAGE_ID_PREFIX.length) : value
  return /^\d+$/.test(id) ? id : null
}
