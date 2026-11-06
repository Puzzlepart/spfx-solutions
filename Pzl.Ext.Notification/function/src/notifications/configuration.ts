import { Settings } from '../config'
import { SharePointClient } from '../sharepoint'

/** One row in the configuration list (`Templates/NotificationConfig.xml`) */
export interface SiteConfiguration {
  id: number
  /** Row title; used in the log, and as the notification source if the site title cannot be read */
  title: string
  /**
   * Shown under notifications as their source: the site's own title (web title),
   * resolved by `resolveSourceName`. Falls back to `title`.
   */
  sourceName: string
  siteUrl: string
  notifyNews: boolean
  notifyAnnouncements: boolean
  /**
   * Severity values (`PzlSeverity`) that are notified without `PzlNotifyTeams`
   * (`PzlNotifySeverities`). Empty means announcements are notified only when
   * `PzlNotifyTeams` is set.
   */
  notifySeverities: string[]
  defaultGroupIds: string[]
  announcementListTitle: string
}

interface ConfigurationItem {
  Id: number
  Title: string
  PzlSiteUrl: string
  PzlNotifyNews: boolean | null
  PzlNotifyAnnouncements: boolean | null
  /** Multi-choice; missing on lists provisioned before the column was added */
  PzlNotifySeverities?: string[] | null
  PzlDefaultGroupIds: string | null
  PzlAnnouncementListTitle: string | null
}

const GUID = /[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/gi

/** Reads the active rows from the configuration list */
export async function getSiteConfigurations(settings: Settings): Promise<SiteConfiguration[]> {
  const sp = new SharePointClient(settings.configSiteUrl)
  const items = await sp.getItems<ConfigurationItem>(
    SharePointClient.listByTitle(settings.configListTitle),
    {
      select: [
        'Id',
        'Title',
        'PzlSiteUrl',
        'PzlNotifyNews',
        'PzlNotifyAnnouncements',
        'PzlDefaultGroupIds',
        'PzlAnnouncementListTitle'
      ],
      // Added after 1.1.0; dropped (and the query retried) when the list does not have it
      optionalSelect: ['PzlNotifySeverities'],
      filter: 'PzlActive eq 1'
    }
  )
  return items
    .filter((item) => item.PzlSiteUrl)
    .map((item) => ({
      id: item.Id,
      title: item.Title,
      sourceName: item.Title,
      siteUrl: item.PzlSiteUrl.trim().replace(/\/+$/, ''),
      notifyNews: item.PzlNotifyNews === true,
      notifyAnnouncements: item.PzlNotifyAnnouncements === true,
      notifySeverities: (item.PzlNotifySeverities ?? []).map((s) => s.trim()).filter(Boolean),
      defaultGroupIds:
        (item.PzlDefaultGroupIds ?? '').match(GUID)?.map((g) => g.toLowerCase()) ?? [],
      announcementListTitle: item.PzlAnnouncementListTitle?.trim() || settings.announcementListTitle
    }))
}

/**
 * Sets `sourceName` to the site's title (`web/Title`), so the line under the
 * notification shows the site as the user knows it rather than the name of the
 * configuration row. Keeps the row title when the site title is empty or cannot
 * be read (the error is logged, the run continues).
 */
export async function resolveSourceName(
  site: SiteConfiguration,
  log: { warn: (message: string) => void }
): Promise<void> {
  try {
    const sp = new SharePointClient(site.siteUrl)
    const web = await sp.get<{ Title: string | null }>('web?$select=Title')
    const title = (web.Title ?? '').trim()
    if (title) site.sourceName = title
  } catch (error) {
    log.warn(`${site.title}: could not read the site title, using the row title: ${error}`)
  }
}
