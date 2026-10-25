import { InvocationContext } from '@azure/functions'
import { Settings } from '../config'
import { sendActivityNotification } from '../graph'
import { odataDate, SharePointClient } from '../sharepoint'
import { SiteConfiguration } from './configuration'
import { buildLink, toPlainText } from './deeplink'
import { RecipientResolver } from './recipients'

/** Activity type in the Teams app manifest */
export const ANNOUNCEMENT_ACTIVITY_TYPE = 'announcement'

interface AnnouncementItem {
  Id: number
  Title: string
  PzlSeverity: string | null
  PzlContent: string | null
  PzlNotifyTeams: boolean | null
  PzlNotifiedAt: string | null
  OData__ModernAudienceTargetUserFieldId?: number[] | null
}

/**
 * Sends notifications for announcements that are active (start date passed, end
 * date not passed), not yet notified, and either have the "always notify"
 * severity or `PzlNotifyTeams` set.
 *
 * `PzlNotifiedAt` is written before sending (claim) and cleared again if the
 * send fails, so overlapping runs never send twice and failures are retried.
 */
export async function processAnnouncements(
  site: SiteConfiguration,
  settings: Settings,
  recipients: RecipientResolver,
  context: InvocationContext
): Promise<void> {
  const sp = new SharePointClient(site.siteUrl)
  const list = SharePointClient.listByTitle(site.announcementListTitle)
  const now = odataDate(new Date())

  const items = await sp.getItems<AnnouncementItem>(list, {
    select: ['Id', 'Title', 'PzlSeverity', 'PzlContent', 'PzlNotifyTeams', 'PzlNotifiedAt'],
    optionalSelect: ['OData__ModernAudienceTargetUserFieldId'],
    filter: `PzlStartDate le ${now} and PzlEndDate ge ${now} and PzlNotifiedAt eq null`
  })

  const candidates = items.filter(
    (item) =>
      item.PzlNotifyTeams === true ||
      settings.alwaysNotifySeverities.includes((item.PzlSeverity ?? '').trim())
  )
  context.log(`${site.title}: ${items.length} active announcements, ${candidates.length} to notify`)

  for (const item of candidates) {
    const label = `announcement ${item.Id} "${item.Title}" on ${site.siteUrl}`
    try {
      const { userIds, groupIds, source } = await recipients.resolve(
        sp,
        item.OData__ModernAudienceTargetUserFieldId ?? undefined,
        site.defaultGroupIds
      )
      if (!userIds.length) {
        context.warn(
          `${label}: no recipients (${source} groups: ${groupIds.join(', ') || 'none'}), skipping`
        )
        continue
      }

      const notification = {
        activityType: ANNOUNCEMENT_ACTIVITY_TYPE,
        source: site.title,
        previewText: toPlainText(item.PzlContent) || item.Title,
        webUrl: buildLink(
          settings,
          `driftsmelding:${item.Id}`,
          `${site.siteUrl}/?driftsmelding=${item.Id}`
        ),
        // Must match the placeholders in the manifest's templateText exactly: Graph's bulk
        // endpoint accepts extra parameters but then silently drops the notification.
        templateParameters: { title: item.Title }
      }

      if (settings.dryRun) {
        context.log(`[dry run] ${label}: would notify ${userIds.length} users (${source})`)
        continue
      }

      // Claim first so an overlapping run skips the item
      await sp.updateItem(list, item.Id, { PzlNotifiedAt: new Date().toISOString() })
      try {
        const count = await sendActivityNotification(notification, userIds, context)
        context.log(`${label}: notified ${count} users (${source} groups: ${groupIds.join(', ')})`)
      } catch (error) {
        await sp.updateItem(list, item.Id, { PzlNotifiedAt: null })
        throw error
      }
    } catch (error) {
      context.error(`${label}: ${error}`)
    }
  }
}
