import { InvocationContext } from '@azure/functions'
import { Settings } from '../config'
import { sendActivityNotification } from '../graph'
import { odataDate, SharePointClient } from '../sharepoint'
import { NotificationState } from '../state'
import { SiteConfiguration } from './configuration'
import { buildLink, toPlainText } from './deeplink'
import { RecipientResolver } from './recipients'

/** Activity type in the Teams app manifest */
export const NEWS_ACTIVITY_TYPE = 'news'

interface PageItem {
  Id: number
  Title: string
  Description: string | null
  FileRef: string
  FirstPublishedDate: string | null
  PromotedState: number
  OData__ModerationStatus: number | null
  OData__UIVersionString: string
  OData__ModernAudienceTargetUserFieldId?: number[] | null
}

/**
 * Sends notifications for news pages (`PromotedState = 2`) that were first
 * published within the lookback window, are published (major version, approved)
 * and have not been notified before. State is kept in Table Storage, because
 * writing to a published page would create a new version.
 */
export async function processNews(
  site: SiteConfiguration,
  settings: Settings,
  recipients: RecipientResolver,
  state: NotificationState,
  context: InvocationContext
): Promise<void> {
  const sp = new SharePointClient(site.siteUrl)
  const pagesUrl = `${sp.serverRelativeUrl === '/' ? '' : sp.serverRelativeUrl}/SitePages`
  const list = SharePointClient.listByUrl(pagesUrl)
  const since = new Date(Date.now() - settings.newsLookbackDays * 24 * 60 * 60 * 1000)

  const items = await sp.getItems<PageItem>(list, {
    select: [
      'Id',
      'Title',
      'Description',
      'FileRef',
      'FirstPublishedDate',
      'PromotedState',
      'OData__ModerationStatus',
      'OData__UIVersionString'
    ],
    optionalSelect: ['OData__ModernAudienceTargetUserFieldId'],
    filter: `PromotedState eq 2 and FirstPublishedDate ge ${odataDate(since)}`,
    orderBy: 'FirstPublishedDate asc'
  })

  const published = items.filter(
    (page) =>
      page.FirstPublishedDate &&
      (page.OData__ModerationStatus ?? 0) === 0 &&
      /\.0$/.test(page.OData__UIVersionString ?? '1.0')
  )
  context.log(
    `${site.title}: ${items.length} news pages since ${since.toISOString()}, ${published.length} published`
  )

  for (const page of published) {
    const label = `news ${page.Id} "${page.Title}" on ${site.siteUrl}`
    const rowKey = `page-${page.Id}`
    try {
      const { userIds, groupIds, source } = await recipients.resolve(
        sp,
        page.OData__ModernAudienceTargetUserFieldId ?? undefined,
        site.defaultGroupIds
      )
      const pageUrl = `${new URL(site.siteUrl).origin}${page.FileRef}`
      const notification = {
        activityType: NEWS_ACTIVITY_TYPE,
        source: site.title,
        previewText: toPlainText(page.Description) || page.Title,
        webUrl: buildLink(settings, `page:${page.FileRef}`, pageUrl),
        templateParameters: { title: page.Title }
      }

      // Claim first; returns false when already sent or being sent
      const claimed = await state.claim(site.siteUrl, rowKey, {
        title: page.Title,
        fileRef: page.FileRef
      })
      if (!claimed) continue

      if (!userIds.length) {
        context.warn(
          `${label}: no recipients (${source} groups: ${groupIds.join(', ') || 'none'}), marked as handled`
        )
        await state.markSent(site.siteUrl, rowKey, 0)
        continue
      }
      if (settings.dryRun) {
        context.log(`[dry run] ${label}: would notify ${userIds.length} users (${source})`)
        continue
      }
      try {
        const count = await sendActivityNotification(notification, userIds, context)
        await state.markSent(site.siteUrl, rowKey, count)
        context.log(`${label}: notified ${count} users (${source} groups: ${groupIds.join(', ')})`)
      } catch (error) {
        await state.release(site.siteUrl, rowKey)
        throw error
      }
    } catch (error) {
      context.error(`${label}: ${error}`)
    }
  }
}
