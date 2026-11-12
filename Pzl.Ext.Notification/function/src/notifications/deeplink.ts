import { Settings } from '../config'

/**
 * Builds the Teams deep link a notification opens.
 *
 * Graph only accepts `webUrl` values on a Teams domain followed by `/l/`, so the
 * link always targets the tab given by `TEAMS_DEEPLINK_APP_ID` and
 * `TEAMS_DEEPLINK_ENTITY_ID`, with `subEntityId` set so the page can focus the
 * item (the Announcement web part reads `driftsmelding:<id>`). `webUrl` is the
 * browser fallback when the tab cannot be opened.
 */
export function buildLink(settings: Settings, subEntityId: string, webUrl: string): string {
  if (!settings.teamsDeepLinkAppId || !settings.teamsDeepLinkEntityId) {
    throw new Error('TEAMS_DEEPLINK_APP_ID and TEAMS_DEEPLINK_ENTITY_ID must be set')
  }
  const context = encodeURIComponent(JSON.stringify({ subEntityId }))
  return `https://teams.microsoft.com/l/entity/${settings.teamsDeepLinkAppId}/${encodeURIComponent(
    settings.teamsDeepLinkEntityId
  )}?webUrl=${encodeURIComponent(webUrl)}&context=${context}`
}

/** Strips HTML tags and collapses whitespace, for preview text */
export function toPlainText(html: string | null | undefined, maxLength = 150): string {
  const text = (html ?? '')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/\s+/g, ' ')
    .trim()
  return text.length > maxLength ? `${text.substring(0, maxLength - 1)}…` : text
}
