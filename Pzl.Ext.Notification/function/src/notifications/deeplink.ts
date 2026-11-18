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

/** Named character references that occur in rich text from SharePoint (the rest are numeric) */
const NAMED_ENTITIES: Record<string, string> = {
  nbsp: ' ',
  amp: '&',
  lt: '<',
  gt: '>',
  quot: '"',
  apos: "'",
  aelig: 'æ',
  AElig: 'Æ',
  oslash: 'ø',
  Oslash: 'Ø',
  aring: 'å',
  Aring: 'Å',
  eacute: 'é',
  ndash: '–',
  mdash: '—',
  hellip: '…',
  laquo: '«',
  raquo: '»'
}

/**
 * Decodes character references (`&#58;`, `&#xE6;`, `&amp;` …). SharePoint
 * stores rich text with such references, also for plain characters like the
 * colon. Unknown names are left as they are.
 */
export function decodeEntities(text: string): string {
  return text.replace(/&(#x[0-9a-f]+|#\d+|[a-z]+);/gi, (match, name: string) => {
    if (name[0] !== '#') return NAMED_ENTITIES[name] ?? match
    const code = /^#x/i.test(name) ? parseInt(name.slice(2), 16) : parseInt(name.slice(1), 10)
    return code > 0 && code <= 0x10ffff ? String.fromCodePoint(code) : match
  })
}

/** Strips HTML tags, decodes character references and collapses whitespace, for preview text */
export function toPlainText(html: string | null | undefined, maxLength = 150): string {
  const text = decodeEntities((html ?? '').replace(/<[^>]+>/g, ' '))
    .replace(/[\u200b\ufeff]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
  return text.length > maxLength ? `${text.substring(0, maxLength - 1)}…` : text
}
