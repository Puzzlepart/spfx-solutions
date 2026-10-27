import { WebPartContext } from '@microsoft/sp-webpart-base'
import { parseAnnouncementId, QUERY_PARAM_NAME } from '../../shared/deepLink'

export {
  SUB_PAGE_ID_PREFIX,
  QUERY_PARAM_NAME,
  PAGE_PREFIX,
  getPageDeepLink,
  getPageRedirectUrl,
  parseAnnouncementId
} from '../../shared/deepLink'

/**
 * Reads `subPageId` from the Teams context. Uses `teamsJs.app.getContext()`
 * (Teams JS v2) and falls back to the deprecated `context.subEntityId`.
 * Returns `null` when not running in Teams, or when the SDK does not answer
 * within `TEAMS_CONTEXT_TIMEOUT_MS` (it can hang in some hosting scenarios).
 */
const TEAMS_CONTEXT_TIMEOUT_MS = 3000

async function getTeamsSubPageId(context: WebPartContext): Promise<string | null> {
  const teams: any = context?.sdks?.microsoftTeams
  if (!teams) return null
  try {
    const app = teams.teamsJs?.app
    if (app?.getContext) {
      const timeout = new Promise<null>((resolve) =>
        setTimeout(() => resolve(null), TEAMS_CONTEXT_TIMEOUT_MS)
      )
      const teamsContext = await Promise.race([app.getContext(), timeout])
      const subPageId = teamsContext?.page?.subPageId
      if (subPageId) return subPageId
    }
  } catch {
    // Fall through to the legacy context below
  }
  return teams.context?.subEntityId ?? null
}

/**
 * Reads the `driftsmelding` query parameter from the page URL. The Teams app's
 * static tab passes `{subEntityId}` this way (see notification/teams-app), so
 * it also carries the deep link value, prefix included.
 */
function getUrlAnnouncementId(): string | null {
  try {
    const params = new URLSearchParams(window.location.search)
    return parseAnnouncementId(params.get(QUERY_PARAM_NAME))
  } catch {
    return null
  }
}

/**
 * Resolves the announcement ID that should be focused (popover opened) on load.
 * The `driftsmelding` query parameter is checked first (synchronous, and what
 * the Teams tab passes), then the Teams deep link context (`subPageId`).
 *
 * @param context Web part context
 */
export async function getFocusedAnnouncementId(context: WebPartContext): Promise<string | null> {
  const fromUrl = getUrlAnnouncementId()
  if (fromUrl) return fromUrl
  return parseAnnouncementId(await getTeamsSubPageId(context))
}
