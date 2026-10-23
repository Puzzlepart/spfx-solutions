import { WebPartContext } from '@microsoft/sp-webpart-base'

/**
 * Prefix used in the Teams deep link `subEntityId`/`subPageId`
 * (e.g. `driftsmelding:12`).
 */
export const SUB_PAGE_ID_PREFIX = 'driftsmelding:'

/**
 * Query string parameter used as fallback outside Teams
 * (e.g. `?driftsmelding=12`).
 */
export const QUERY_PARAM_NAME = 'driftsmelding'

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

/**
 * Reads `subPageId` from the Teams context. Uses `teamsJs.app.getContext()`
 * (Teams JS v2) and falls back to the deprecated `context.subEntityId`.
 * Returns `null` when not running in Teams.
 */
async function getTeamsSubPageId(context: WebPartContext): Promise<string | null> {
  const teams: any = context?.sdks?.microsoftTeams
  if (!teams) return null
  try {
    const app = teams.teamsJs?.app
    if (app?.getContext) {
      const teamsContext = await app.getContext()
      const subPageId = teamsContext?.page?.subPageId
      if (subPageId) return subPageId
    }
  } catch {
    // Fall through to the legacy context below
  }
  return teams.context?.subEntityId ?? null
}

/**
 * Resolves the announcement ID that should be focused (popover opened) on load.
 * Teams deep link (`subPageId`) takes precedence over the `driftsmelding` query parameter.
 *
 * @param context Web part context
 */
export async function getFocusedAnnouncementId(context: WebPartContext): Promise<string | null> {
  const fromTeams = parseAnnouncementId(await getTeamsSubPageId(context))
  if (fromTeams) return fromTeams
  try {
    const params = new URLSearchParams(window.location.search)
    return parseAnnouncementId(params.get(QUERY_PARAM_NAME))
  } catch {
    return null
  }
}
