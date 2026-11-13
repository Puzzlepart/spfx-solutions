import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions'
import { getSettings } from '../config'
import { getAccessToken } from '../auth'
import { HttpError, request } from '../http'
import { buildLink } from '../notifications/deeplink'

const GRAPH = 'https://graph.microsoft.com/v1.0'

/**
 * POST /api/test-notification?userId=<Entra user id>[&type=announcement|news][&bulk=1][&extra=1]
 *        [&teamsAppId=<id>][&installationId=<id>]
 *
 * Sends one activity notification to a single user. By default it uses the
 * per-user Graph endpoint, which returns delivery errors synchronously, unlike
 * the bulk endpoint the timer uses. Use it to verify the Teams app, permissions
 * and deep link for one recipient. Protected by a function key.
 *
 * bulk=1  uses sendActivityNotificationToRecipients (the timer's endpoint)
 * extra=1 adds a template parameter that is not in the manifest's templateText
 * teamsAppId=<id> sends that Teams app ID to Graph (default: setting TEAMS_NOTIFICATION_APP_ID),
 *         to try a value before it is stored in the setting
 * installationId=<id> uses the user's installation of the Teams app as the topic instead of a
 *         text topic with a deep link (id from GET /users/{id}/teamwork/installedApps). Tells
 *         whether a 403 is about the app installation or about the text topic and its link.
 *
 * A failed call also returns the claims of the Graph token that decide what it may do (app ID,
 * tenant, roles), never the token itself.
 */
export async function testNotification(
  req: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  const settings = getSettings()
  const userId = req.query.get('userId')
  const type = req.query.get('type') === 'news' ? 'news' : 'announcement'
  const bulk = req.query.get('bulk') === '1'
  const extra = req.query.get('extra') === '1'
  const teamsAppId = req.query.get('teamsAppId') || settings.notificationTeamsAppId
  const installationId = req.query.get('installationId')
  if (!userId) return { status: 400, jsonBody: { error: 'userId query parameter is required' } }

  // Both types open the tab. news: `page:<server-relative path>` of pageUrl (defaults to the
  // config site's home page), redirected by the NotificationDeepLink customizer; announcement:
  // focus on a (non-existing) item
  const pageUrl = req.query.get('pageUrl') || `${settings.configSiteUrl}/SitePages/Home.aspx`
  const webUrl =
    type === 'news'
      ? buildLink(settings, `page:${new URL(pageUrl).pathname}`, pageUrl)
      : buildLink(settings, 'driftsmelding:0', settings.configSiteUrl)
  const stamp = new Date().toISOString()
  const body: Record<string, unknown> = {
    topic: installationId
      ? {
          source: 'entityUrl',
          value: `${GRAPH}/users/${userId}/teamwork/installedApps/${installationId}`
        }
      : { source: 'text', value: new URL(settings.configSiteUrl).host, webUrl },
    activityType: type,
    previewText: {
      content: `Testvarsel (${bulk ? 'bulk' : 'single'}${extra ? ', extra' : ''}) ${stamp}`
    },
    templateParameters: [
      { name: 'title', value: 'Testvarsel' },
      ...(extra ? [{ name: 'severity', value: 'Varsel' }] : [])
    ]
  }
  if (teamsAppId) body.teamsAppId = teamsAppId
  if (bulk) {
    body.recipients = [{ '@odata.type': 'microsoft.graph.aadUserNotificationRecipient', userId }]
  }
  const url = bulk
    ? `${GRAPH}/teamwork/sendActivityNotificationToRecipients`
    : `${GRAPH}/users/${userId}/teamwork/sendActivityNotification`
  context.log(
    `Test notification (${type}, bulk=${bulk}, extra=${extra}) to ${userId}: ${JSON.stringify(body)}`
  )

  const result = { userId, type, bulk, extra, teamsAppId, installationId, webUrl, stamp }
  let token = ''
  try {
    token = await getAccessToken('https://graph.microsoft.com/.default')
    await request(url, {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}` },
      body,
      retries: 0
    })
    return { status: 200, jsonBody: { ok: true, ...result } }
  } catch (error) {
    const status = error instanceof HttpError ? error.status : 500
    const detail = error instanceof HttpError ? safeJson(error.body) : String(error)
    context.error(`Test notification to ${userId} failed: ${status} ${JSON.stringify(detail)}`)
    return {
      status: 200,
      jsonBody: {
        ok: false,
        ...result,
        graphStatus: status,
        graphError: detail,
        token: tokenClaims(token)
      }
    }
  }
}

/** The claims of an access token that decide what it may do. Never returns the token itself. */
function tokenClaims(token: string): Record<string, unknown> {
  try {
    const claims = JSON.parse(Buffer.from(token.split('.')[1], 'base64url').toString('utf8'))
    return {
      appid: claims.appid ?? claims.azp,
      tid: claims.tid,
      idtyp: claims.idtyp,
      roles: claims.roles ?? []
    }
  } catch {
    return {}
  }
}

function safeJson(text: string): unknown {
  try {
    return JSON.parse(text)
  } catch {
    return text
  }
}

app.http('testNotification', {
  route: 'test-notification',
  methods: ['POST'],
  authLevel: 'function',
  handler: testNotification
})
