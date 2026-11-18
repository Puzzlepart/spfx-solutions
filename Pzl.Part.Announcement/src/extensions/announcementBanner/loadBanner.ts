import { BaseComponentContext } from '@microsoft/sp-component-base'
import strings from 'AnnouncementStrings'
import { AnnouncementListService } from '../../Announcement/services/AnnouncementListService'
import { fetchAnnouncements } from '../../Announcement/services/fetchAnnouncements'
import { DEFAULT_POPOVER_FIELDS, NOTIFY_TEAMS_FIELD } from '../../Announcement/constants'
import type { IAnnouncement } from '../../Announcement/components/types'
import type { IAnnouncementBannerProps } from './components/AnnouncementBanner'

export interface IBannerListOptions {
  listSiteUrl?: string
  listTitle?: string
  targetAudience?: boolean
  popoverFields?: Record<string, boolean>
}

/**
 * Reads the announcement list and returns the props for the banner, or
 * `null` when there is no active announcement marked "Vis som banner".
 * For users who may edit the list it also returns what the editor functions
 * of the dialog need. Loaded on demand by the customizer, after the targeting check, so the
 * PnPjs code is fetched only on pages that may show a banner.
 */
export async function loadBanner(
  context: BaseComponentContext,
  options: IBannerListOptions
): Promise<IAnnouncementBannerProps | null> {
  const service = new AnnouncementListService(
    context,
    options.listTitle?.trim() || strings.AnnouncementsListName,
    options.listSiteUrl?.trim() || undefined
  )
  const fields = await service.getFields()
  const load = async (): Promise<IAnnouncement[]> =>
    (await fetchAnnouncements(service, context, fields, options.targetAudience === true)).filter(
      (announcement) => announcement.showAsBanner
    )
  const announcements = await load()
  if (!announcements.length) return null

  const popoverFields = options.popoverFields ?? DEFAULT_POPOVER_FIELDS
  const detailFields = fields.filter((field) => popoverFields[field.internalName] === true)

  // Editors get the same functions in the dialog as in the web part
  const canEdit = await service.canEditItems()
  const editor = canEdit
    ? {
        service,
        fields,
        alwaysNotified: fields.some((field) => field.internalName === NOTIFY_TEAMS_FIELD)
          ? await service.getAlwaysNotifiedSeverities()
          : []
      }
    : undefined
  return { announcements, fields: detailFields, editor, reload: load }
}
