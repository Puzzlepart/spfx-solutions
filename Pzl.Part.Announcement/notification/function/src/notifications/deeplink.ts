import { Settings } from '../config'

/**
 * Builds the URL a notification opens.
 *
 * When `TEAMS_DEEPLINK_APP_ID` and `TEAMS_DEEPLINK_ENTITY_ID` are set, the link
 * opens that tab in Teams with `subEntityId` set, so the page can focus the
 * item (the Announcement web part reads `driftsmelding:<id>`). Otherwise the
 * SharePoint URL is used directly.
 */
export function buildLink(settings: Settings, subEntityId: string, fallbackUrl: string): string {
  if (!settings.teamsDeepLinkAppId || !settings.teamsDeepLinkEntityId) return fallbackUrl
  const context = encodeURIComponent(JSON.stringify({ subEntityId }))
  return `https://teams.microsoft.com/l/entity/${settings.teamsDeepLinkAppId}/${encodeURIComponent(
    settings.teamsDeepLinkEntityId
  )}?context=${context}`
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
