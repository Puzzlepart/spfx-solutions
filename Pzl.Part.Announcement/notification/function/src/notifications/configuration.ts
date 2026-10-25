import { Settings } from '../config'
import { SharePointClient } from '../sharepoint'

/** One row in the configuration list (`Templates/NotificationConfig.xml`) */
export interface SiteConfiguration {
  id: number
  /** Row title; shown under notifications as their source (e.g. the intranet name) */
  title: string
  siteUrl: string
  notifyNews: boolean
  notifyAnnouncements: boolean
  defaultGroupIds: string[]
  announcementListTitle: string
}

interface ConfigurationItem {
  Id: number
  Title: string
  PzlSiteUrl: string
  PzlNotifyNews: boolean | null
  PzlNotifyAnnouncements: boolean | null
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
      filter: 'PzlActive eq 1'
    }
  )
  return items
    .filter((item) => item.PzlSiteUrl)
    .map((item) => ({
      id: item.Id,
      title: item.Title,
      siteUrl: item.PzlSiteUrl.trim().replace(/\/+$/, ''),
      notifyNews: item.PzlNotifyNews === true,
      notifyAnnouncements: item.PzlNotifyAnnouncements === true,
      defaultGroupIds:
        (item.PzlDefaultGroupIds ?? '').match(GUID)?.map((g) => g.toLowerCase()) ?? [],
      announcementListTitle: item.PzlAnnouncementListTitle?.trim() || settings.announcementListTitle
    }))
}
