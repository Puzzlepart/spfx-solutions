import { InvocationContext } from '@azure/functions'
import { Settings } from '../config'
import { sendActivityNotification } from '../graph'
import { odataDate, SharePointClient } from '../sharepoint'
import { BEFORE_START_NOTE, NotificationState } from '../state'
import { SiteConfiguration } from './configuration'
import { buildLink, toPlainText } from './deeplink'
import { Status, toSeverity, toStatus } from './levels'
import { RecipientResolver } from './recipients'

/** Activity type in the Teams app manifest */
export const ANNOUNCEMENT_ACTIVITY_TYPE = 'announcement'

interface AnnouncementItem {
  Id: number
  Title: string
  PzlSeverity: string | null
  PzlContent: string | null
  PzlNotifiedAt: string | null
  Created: string
  PzlStartDate: string
  /** Optional: lists provisioned before Driftsmeldinger 2.0 do not have it */
  PzlStatus?: string | null
  /** Optional: added by `Templates/AnnouncementFields.xml` */
  PzlNotifyTeams?: boolean | null
  /** Optional: added by `Templates/AnnouncementFields.xml` (2.0) */
  PzlNotifyAgain?: boolean | null
  OData__ModernAudienceTargetUserFieldId?: number[] | null
}

/**
 * Why an announcement is notified: the first notification when it becomes
 * active, or a new one requested with «Varsle på nytt»
 */
type NotificationKind = 'first' | 'again'

/**
 * Title shown in the bell (`{title}` in the manifest's `templateText`):
 * the first notification says what kind of announcement it is, a new one
 * says that it was updated or resolved.
 */
export function notificationTitle(kind: NotificationKind, status: Status, title: string): string {
  if (kind === 'again') return status === 'resolved' ? `Løst: ${title}` : `Oppdatert: ${title}`
  return status === 'planned' ? `Planlagt: ${title}` : `Driftsmelding: ${title}`
}

/**
 * Decides whether an active announcement is notified in this run, and how.
 *
 * - «Varsle på nytt» (`PzlNotifyAgain`) always sends a new notification.
 * - Otherwise the first notification is sent once (`PzlNotifiedAt` empty), when
 *   «Varsle i Teams» is set or the level is one of the site's «Nivåer som
 *   alltid varsles», unless the announcement is already resolved.
 */
export function classify(item: AnnouncementItem, site: SiteConfiguration): NotificationKind | null {
  if (item.PzlNotifyAgain === true) return 'again'
  if (item.PzlNotifiedAt) return null
  if (toStatus(item.PzlStatus) === 'resolved') return null
  const levelNotifies = site.notifySeverities.includes(toSeverity(item.PzlSeverity))
  return item.PzlNotifyTeams === true || levelNotifies ? 'first' : null
}

/**
 * Whether an announcement is from before the notifications started for the site:
 * it existed then, and its start date had passed. Such an announcement gets no
 * first notification, so turning the service on does not notify about what is
 * already there. One that is created later, or that becomes active later, is new.
 */
export function isFromBeforeStart(
  item: Pick<AnnouncementItem, 'Created' | 'PzlStartDate'>,
  start: Date
): boolean {
  return new Date(item.Created) < start && new Date(item.PzlStartDate) < start
}

/**
 * Sends notifications for active announcements (start date passed, end date
 * not passed), see `classify`. A first notification is only sent for
 * announcements that are new since the notifications started for the site
 * (`isFromBeforeStart`); «Varsle på nytt» always sends.
 *
 * The item is claimed before sending (`PzlNotifiedAt` set, `PzlNotifyAgain`
 * cleared) and the claim is undone if the send fails, so overlapping runs never
 * send twice and failures are retried.
 */
export async function processAnnouncements(
  site: SiteConfiguration,
  settings: Settings,
  recipients: RecipientResolver,
  state: NotificationState,
  context: InvocationContext
): Promise<void> {
  const sp = new SharePointClient(site.siteUrl)
  const list = SharePointClient.listByTitle(site.announcementListTitle)
  const now = odataDate(new Date())

  // Filtered on dates only: the optional columns cannot be used in $filter on
  // lists that do not have them, and there are few active announcements.
  const items = await sp.getItems<AnnouncementItem>(list, {
    select: [
      'Id',
      'Title',
      'PzlSeverity',
      'PzlContent',
      'PzlNotifiedAt',
      'Created',
      'PzlStartDate'
    ],
    optionalSelect: [
      'PzlStatus',
      'PzlNotifyTeams',
      'PzlNotifyAgain',
      'OData__ModernAudienceTargetUserFieldId'
    ],
    filter: `PzlStartDate le ${now} and PzlEndDate ge ${now}`
  })

  // A dry run before the start shows what is there, to check the recipients, and says that it
  // will not be sent; after the start, what is from before it is left out.
  const start = await state.notificationStart(site.siteUrl, 'announcement')
  const candidates: { item: AnnouncementItem; kind: NotificationKind; old: boolean }[] = []
  let fromBefore = 0
  for (const item of items) {
    const kind = classify(item, site)
    if (!kind) continue
    const old = kind === 'first' && isFromBeforeStart(item, start.at)
    if (old) fromBefore++
    if (!old || !start.stored) candidates.push({ item, kind, old })
  }
  context.log(
    `${site.title}: ${items.length} active announcements, ${candidates.filter((candidate) => !candidate.old).length} to notify, ${fromBefore} from before the notifications started (levels always notified: ${site.notifySeverities.join(', ') || 'none'})`
  )

  for (const { item, kind, old } of candidates) {
    const label = `announcement ${item.Id} "${item.Title}" (${kind}) on ${site.siteUrl}`
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
        source: site.sourceName,
        previewText: toPlainText(item.PzlContent) || item.Title,
        webUrl: buildLink(
          settings,
          `driftsmelding:${item.Id}`,
          `${site.siteUrl}/?driftsmelding=${item.Id}`
        ),
        // Must match the placeholders in the manifest's templateText exactly: Graph's bulk
        // endpoint accepts extra parameters but then silently drops the notification.
        templateParameters: {
          title: notificationTitle(kind, toStatus(item.PzlStatus), item.Title)
        }
      }

      if (settings.dryRun) {
        context.log(
          old
            ? `[dry run] ${label}: ${userIds.length} recipients (${source}). ${BEFORE_START_NOTE}`
            : `[dry run] ${label}: would notify ${userIds.length} users (${source})`
        )
        continue
      }

      // Claim first so an overlapping run skips the item
      const claim: Record<string, unknown> = { PzlNotifiedAt: new Date().toISOString() }
      const undo: Record<string, unknown> = { PzlNotifiedAt: item.PzlNotifiedAt ?? null }
      if (kind === 'again') {
        claim.PzlNotifyAgain = false
        undo.PzlNotifyAgain = true
      }
      await sp.updateItem(list, item.Id, claim)
      try {
        const count = await sendActivityNotification(
          notification,
          userIds,
          context,
          settings.notificationTeamsAppId
        )
        context.log(`${label}: notified ${count} users (${source} groups: ${groupIds.join(', ')})`)
      } catch (error) {
        await sp.updateItem(list, item.Id, undo)
        throw error
      }
    } catch (error) {
      context.error(`${label}: ${error}`)
    }
  }
}
