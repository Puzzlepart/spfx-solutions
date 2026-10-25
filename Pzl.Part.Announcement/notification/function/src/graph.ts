import { InvocationContext } from '@azure/functions'
import { getAccessToken } from './auth'
import { request } from './http'

const GRAPH = 'https://graph.microsoft.com/v1.0'
const GRAPH_SCOPE = 'https://graph.microsoft.com/.default'
/** Maximum recipients per sendActivityNotificationToRecipients call */
const RECIPIENT_BATCH_SIZE = 100

async function graphHeaders(): Promise<Record<string, string>> {
  return { Authorization: `Bearer ${await getAccessToken(GRAPH_SCOPE)}` }
}

interface GraphUser {
  id: string
  accountEnabled?: boolean
  userType?: string
}

/**
 * Returns the user IDs of all transitive members of a group. Disabled accounts are
 * skipped; guests are skipped unless `includeGuests` is set.
 */
export async function getGroupMemberIds(
  groupId: string,
  includeGuests: boolean
): Promise<string[]> {
  const ids: string[] = []
  let url: string | undefined =
    `${GRAPH}/groups/${groupId}/transitiveMembers/microsoft.graph.user?$select=id,accountEnabled,userType&$top=999`
  while (url) {
    const page: { value: GraphUser[]; '@odata.nextLink'?: string } = await request(url, {
      headers: await graphHeaders()
    })
    for (const user of page.value) {
      if (user.accountEnabled === false) continue
      if (!includeGuests && user.userType === 'Guest') continue
      ids.push(user.id)
    }
    url = page['@odata.nextLink']
  }
  return ids
}

export interface ActivityNotification {
  /** Must match an `activityType` in the Teams app manifest */
  activityType: string
  /** Shown as the notification title */
  title: string
  /** Shown as the notification body (max 150 characters are shown) */
  previewText: string
  /** Where a click on the notification goes */
  webUrl: string
  /**
   * Values for the `templateText` placeholders in the manifest. Send exactly the
   * placeholders the manifest declares: the bulk endpoint accepts unknown parameters
   * but then drops the notification without an error (the per-user endpoint returns
   * "Incorrect template parameter arity").
   */
  templateParameters?: Record<string, string>
}

/**
 * Sends an activity notification to a set of users, in batches of 100.
 * Returns the number of recipients the notification was sent to.
 */
export async function sendActivityNotification(
  notification: ActivityNotification,
  userIds: string[],
  context: InvocationContext
): Promise<number> {
  const recipients = Array.from(new Set(userIds))
  for (let i = 0; i < recipients.length; i += RECIPIENT_BATCH_SIZE) {
    const batch = recipients.slice(i, i + RECIPIENT_BATCH_SIZE)
    await request(`${GRAPH}/teamwork/sendActivityNotificationToRecipients`, {
      method: 'POST',
      headers: await graphHeaders(),
      body: {
        topic: {
          source: 'text',
          value: notification.title.substring(0, 100),
          webUrl: notification.webUrl
        },
        activityType: notification.activityType,
        previewText: { content: notification.previewText.substring(0, 150) },
        recipients: batch.map((userId) => ({
          '@odata.type': 'microsoft.graph.aadUserNotificationRecipient',
          userId
        })),
        templateParameters: Object.entries(notification.templateParameters ?? {}).map(
          ([name, value]) => ({ name, value })
        )
      }
    })
    context.log(
      `Sent "${notification.title}" (${notification.activityType}) to ${batch.length} recipients (${i + batch.length}/${recipients.length})`
    )
  }
  return recipients.length
}
