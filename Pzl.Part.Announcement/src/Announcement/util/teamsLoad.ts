/**
 * Teams keeps the tab's page alive (app caching) when the user opens the tab
 * again, for example by clicking another notification in the activity feed: the
 * page is not loaded again, so its address keeps the deep link of the first
 * notification. Teams tells the page with a `load` message that carries the
 * address the tab would have been loaded with:
 *
 * `{ func: 'load', args: [{ entityId, contentUrl }] }`
 *
 * SharePoint answers the message but does not navigate, so the web part listens
 * for it too and reads the deep link from `contentUrl`.
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
 * Returns `contentUrl` of a Teams `load` message, or `null` for anything else
 *
 * @param data Data of the message event
 */
export function getTeamsLoadContentUrl(data: unknown): string | null {
  if (!data || typeof data !== 'object') return null
  const message = data as { func?: unknown; args?: unknown }
  if (message.func !== 'load' || !Array.isArray(message.args)) return null
  const contentUrl = (message.args[0] as { contentUrl?: unknown } | undefined)?.contentUrl
  return typeof contentUrl === 'string' && contentUrl ? contentUrl : null
}

/**
 * Calls `callback` with the tab address each time Teams shows the page again
 * without loading it. Returns a function that stops listening.
 *
 * @param callback Called with `contentUrl` of the `load` message
 */
export function onTeamsLoad(callback: (contentUrl: string) => void): () => void {
  const listener = (event: MessageEvent): void => {
    if (!isTeamsHostOrigin(event.origin)) return
    const contentUrl = getTeamsLoadContentUrl(event.data)
    if (contentUrl) callback(contentUrl)
  }
  window.addEventListener('message', listener)
  return () => window.removeEventListener('message', listener)
}
