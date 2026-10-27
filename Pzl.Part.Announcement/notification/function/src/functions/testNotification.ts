import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions'
import { getSettings } from '../config'
import { getAccessToken } from '../auth'
import { HttpError, request } from '../http'
import { buildLink } from '../notifications/deeplink'

const GRAPH = 'https://graph.microsoft.com/v1.0'

/**
 * POST /api/test-notification?userId=<Entra user id>[&type=announcement|news][&bulk=1][&extra=1]
 *
 * Sends one activity notification to a single user. By default it uses the
 * per-user Graph endpoint, which returns delivery errors synchronously, unlike
 * the bulk endpoint the timer uses. Use it to verify the Teams app, permissions
 * and deep link for one recipient. Protected by a function key.
 *
 * bulk=1  uses sendActivityNotificationToRecipients (the timer's endpoint)
 * extra=1 adds a template parameter that is not in the manifest's templateText
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
  if (!userId) return { status: 400, jsonBody: { error: 'userId query parameter is required' } }

  // Both types open the tab. news: `page:<server-relative path>` of pageUrl (defaults to the
  // config site's home page), redirected by the AnnouncementDeepLink customizer; announcement:
  // focus on a (non-existing) item
  const pageUrl = req.query.get('pageUrl') || `${settings.configSiteUrl}/SitePages/Home.aspx`
  const webUrl =
    type === 'news'
      ? buildLink(settings, `page:${new URL(pageUrl).pathname}`, pageUrl)
      : buildLink(settings, 'driftsmelding:0', settings.configSiteUrl)
  const stamp = new Date().toISOString()
  const body: Record<string, unknown> = {
    topic: { source: 'text', value: new URL(settings.configSiteUrl).host, webUrl },
    activityType: type,
    previewText: {
      content: `Testvarsel (${bulk ? 'bulk' : 'single'}${extra ? ', extra' : ''}) ${stamp}`
    },
    templateParameters: [
      { name: 'title', value: 'Testvarsel' },
      ...(extra ? [{ name: 'severity', value: 'Varsel' }] : [])
    ]
  }
  if (bulk) {
    body.recipients = [{ '@odata.type': 'microsoft.graph.aadUserNotificationRecipient', userId }]
  }
  const url = bulk
    ? `${GRAPH}/teamwork/sendActivityNotificationToRecipients`
    : `${GRAPH}/users/${userId}/teamwork/sendActivityNotification`
  context.log(
    `Test notification (${type}, bulk=${bulk}, extra=${extra}) to ${userId}: ${JSON.stringify(body)}`
  )

  const result = { userId, type, bulk, extra, webUrl, stamp }
  try {
    await request(url, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${await getAccessToken('https://graph.microsoft.com/.default')}`
      },
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
      jsonBody: { ok: false, ...result, graphStatus: status, graphError: detail }
    }
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
