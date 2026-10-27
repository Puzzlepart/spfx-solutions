import { BaseComponentContext } from '@microsoft/sp-component-base'
import { graphfi, SPFx as graphSPFx } from '@pnp/graph'
import '@pnp/graph/users'
import '@pnp/graph/groups'
import '@pnp/graph/members'
import '@pnp/sp/webs'
import '@pnp/sp/lists'
import '@pnp/sp/items'
import '@pnp/sp/site-users/web'
import { MessageBarIntent } from '@fluentui/react-components'
import strings from 'AnnouncementStrings'
import { IAnnouncement, IListField } from '../components/types'
import { OPTIONAL_FIELDS } from '../constants'
import type { AnnouncementListService } from './AnnouncementListService'

/**
 * Builds the `$select`/`$expand` for the list query from the list fields.
 * User and lookup fields are expanded to get title (and e-mail).
 */
function buildQuery(fields: IListField[], targetAudience: boolean) {
  const select = ['ID']
  const expand: string[] = []
  fields.forEach((field) => {
    const name = field.entityPropertyName
    switch (field.type) {
      case 'User':
      case 'UserMulti':
        select.push(`${name}/Id`, `${name}/Title`, `${name}/EMail`)
        expand.push(name)
        break
      case 'Lookup':
      case 'LookupMulti':
        select.push(`${name}/Id`, `${name}/Title`)
        expand.push(name)
        break
      default:
        select.push(name)
    }
  })
  if (targetAudience && !select.includes('OData__ModernAudienceTargetUserFieldId')) {
    select.push('OData__ModernAudienceTargetUserFieldId')
  }
  return { select, expand }
}

/**
 * Maps the `PzlSeverity` choice value to a message bar intent
 */
export function toSeverity(value: string | undefined): MessageBarIntent {
  switch (value) {
    case strings.Severity.Success:
      return 'success'
    case strings.Severity.Warning:
      return 'warning'
    case strings.Severity.Error:
      return 'error'
    default:
      return 'info'
  }
}

/**
 * Fetches the active announcements (start date passed, end date not passed)
 * the current user may see, pinned first. Shared by the web part and the
 * AnnouncementBanner application customizer.
 *
 * @param service List service
 * @param context Component context (for the Graph client)
 * @param fields Fields to select (see `AnnouncementListService.getFields`)
 * @param targetAudience Filter on the list's audience targeting column
 */
export async function fetchAnnouncements(
  service: AnnouncementListService,
  context: BaseComponentContext,
  fields: IListField[],
  targetAudience: boolean
): Promise<IAnnouncement[]> {
  const sp = service.sp

  const fetchItems = (selectFields: IListField[]): Promise<any[]> => {
    const now = new Date().toISOString()
    const { select, expand } = buildQuery(selectFields, targetAudience)
    let query = service.list.items
      .select(...select)
      .filter(`PzlStartDate le datetime'${now}' and PzlEndDate ge datetime'${now}'`)
    if (expand.length > 0) query = query.expand(...expand)
    return query()
  }

  try {
    const graph = graphfi().using(graphSPFx(context))

    let spItems: any[]
    try {
      spItems = await fetchItems(fields)
    } catch {
      // List not updated with the latest template, fall back to the original fields
      spItems = await fetchItems(
        fields.filter((field) => !OPTIONAL_FIELDS.includes(field.internalName))
      )
    }

    const memberOfGroups = targetAudience && (await graph.me.getMemberGroups())

    const announcements = await Promise.all(
      spItems.map(async (item): Promise<IAnnouncement> => {
        const targetUserIds = item.OData__ModernAudienceTargetUserFieldId

        let hasAccess = true
        if (targetAudience && targetUserIds) {
          const groupIds = await Promise.all(
            targetUserIds.map(async (id) => {
              const { LoginName } = await sp.web.getUserById(id)()
              return LoginName.split('|').pop()
            })
          )
          hasAccess = groupIds.some((id) => memberOfGroups.includes(id))
        }

        return {
          id: String(item.ID),
          title: item.Title,
          severity: toSeverity(item.PzlSeverity),
          content: item.PzlContent,
          hasAccess,
          pinned: item.PzlPinned === true,
          showAsBanner: item.PzlShowAsBanner === true,
          notifiedAt: item.PzlNotifiedAt ?? null,
          values: item
        }
      })
    )
    return (
      announcements
        .filter((announcement) => announcement.hasAccess)
        // Pinned first, otherwise keep list order (sort is stable)
        .sort((a, b) => Number(b.pinned) - Number(a.pinned))
    )
  } catch {
    throw new Error(strings.AnnouncementFetchErrorText)
  }
}
