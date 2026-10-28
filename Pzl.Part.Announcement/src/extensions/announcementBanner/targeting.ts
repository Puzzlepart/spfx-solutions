/**
 * Decides whether the banner applies to the current page. Kept free of SPFx
 * imports so the entry bundle stays small and the logic can be tested.
 */

export interface IBannerTargeting {
  /**
   * Sites the banner is shown on, as absolute URLs. A site collection URL
   * also covers its subsites. Case-insensitive, trailing slash ignored.
   */
  siteUrls?: string[]
  /**
   * Hubs whose member sites show the banner: the hub site's absolute URL or
   * its hub site ID (GUID). The hub site itself is a member of its own hub.
   */
  hubSiteUrls?: string[]
  /**
   * Show the banner on every site the customizer is registered on,
   * regardless of `siteUrls` and `hubSiteUrls`. Off by default, so a
   * tenant-wide registration with no targeting shows nothing.
   */
  everywhere?: boolean
}

export interface IPageLocation {
  /** `pageContext.site.absoluteUrl` */
  siteUrl: string
  /** `pageContext.web.absoluteUrl` */
  webUrl: string
  /** URL of the hub the site belongs to, or `null` */
  hubSiteUrl: string | null
  /** ID of the hub the site belongs to, or `null` */
  hubSiteId: string | null
}

const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i

/** Lower case, no trailing slash, no surrounding whitespace */
export function normalizeUrl(url: string | undefined | null): string {
  return (url ?? '').trim().toLowerCase().replace(/\/+$/, '')
}

function list(values: string[] | undefined): string[] {
  return Array.isArray(values) ? values.map(normalizeUrl).filter(Boolean) : []
}

/**
 * Whether `entry` is just an origin (`https://tenant.sharepoint.com`), i.e. the
 * root site collection. It must not act as a prefix for every other site.
 */
function isOriginOnly(entry: string): boolean {
  const afterScheme = entry.indexOf('://') === -1 ? entry : entry.split('://')[1]
  return afterScheme.indexOf('/') === -1
}

/**
 * Whether `url` is `entry` or lives under it (`entry` followed by `/`). The
 * root site collection matches exactly only; its subsites are covered by the
 * site URL check in `isTargeted`.
 */
function isSameOrUnder(url: string, entry: string): boolean {
  if (url === entry) return true
  return !isOriginOnly(entry) && url.indexOf(`${entry}/`) === 0
}

/**
 * Returns `true` when the banner should be shown on the page.
 *
 * @param targeting Targeting properties of the custom action
 * @param page Where the page lives
 */
export function isTargeted(targeting: IBannerTargeting, page: IPageLocation): boolean {
  if (targeting.everywhere === true) return true

  const siteUrl = normalizeUrl(page.siteUrl)
  const webUrl = normalizeUrl(page.webUrl)
  if (list(targeting.siteUrls).some((entry) => isSameOrUnder(webUrl, entry) || siteUrl === entry)) {
    return true
  }

  const hubUrl = normalizeUrl(page.hubSiteUrl)
  const hubId = normalizeUrl(page.hubSiteId)
  return list(targeting.hubSiteUrls).some((entry) =>
    GUID.test(entry) ? hubId === entry : hubUrl !== '' && hubUrl === entry
  )
}

/**
 * Reads the page location from the SPFx page context. `legacyPageContext`
 * carries the hub association (`hubSiteId` and `hubSiteData`, a JSON string
 * with the hub's `url`).
 */
export function getPageLocation(pageContext: {
  site: { absoluteUrl: string }
  web: { absoluteUrl: string }
  legacyPageContext?: any
}): IPageLocation {
  const legacy = pageContext.legacyPageContext ?? {}
  let hubSiteUrl: string | null = null
  try {
    const data = legacy.hubSiteData
    const parsed = typeof data === 'string' ? JSON.parse(data) : data
    hubSiteUrl = parsed?.url ?? null
  } catch {
    hubSiteUrl = null
  }
  const hubSiteId: string | null = legacy.hubSiteId || null
  return {
    siteUrl: pageContext.site.absoluteUrl,
    webUrl: pageContext.web.absoluteUrl,
    hubSiteUrl,
    hubSiteId: hubSiteId && /^0{8}-/.test(hubSiteId) ? null : hubSiteId
  }
}
