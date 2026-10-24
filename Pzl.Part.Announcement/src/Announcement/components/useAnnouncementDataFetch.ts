import { useEffect } from 'react'
import { IAnnouncement, IAnnouncementProps, IAnnouncementState, IListField } from './types'
import { graphfi, SPFx as graphSPFx } from '@pnp/graph'
import '@pnp/graph/users'
import '@pnp/graph/groups'
import '@pnp/graph/members'
import '@pnp/sp/webs'
import '@pnp/sp/lists'
import '@pnp/sp/items'
import '@pnp/sp/site-users/web'
import strings from 'AnnouncementStrings'
import { MessageBarIntent } from '@fluentui/react-components'
import { NOTIFICATION_FIELDS } from '../constants'

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
 * Component data fetch hook for `Announcement`. This hook is responsible for
 * fetching the list schema, the current user's permissions and the active
 * announcements, and setting state.
 *
 * @param props Props
 * @param setState Set state callback
 *
 * @returns `refresh` to fetch the announcements again (e.g. after creating one)
 */
export function useAnnouncementDataFetch(
  props: IAnnouncementProps,
  setState: (newState: Partial<IAnnouncementState>) => void
) {
  const { service } = props
  const sp = service.sp

  /**
   * Fetches active items from the announcement list.
   *
   * @param fields Fields to select
   */
  const fetchItems = (fields: IListField[]): Promise<any[]> => {
    const now = new Date().toISOString()
    const { select, expand } = buildQuery(fields, props.targetAudience)
    let query = service.list.items
      .select(...select)
      .filter(`PzlStartDate le datetime'${now}' and PzlEndDate ge datetime'${now}'`)
    if (expand.length > 0) query = query.expand(...expand)
    return query()
  }

  const getAnnouncements = async (fields: IListField[]): Promise<IAnnouncement[]> => {
    try {
      const graph = graphfi().using(graphSPFx(props.context))

      let spItems: any[]
      try {
        spItems = await fetchItems(fields)
      } catch {
        // List not updated with the v1.1.0 template, fall back to the original fields
        spItems = await fetchItems(
          fields.filter((field) => !NOTIFICATION_FIELDS.includes(field.internalName))
        )
      }

      const memberOfGroups = props.targetAudience && (await graph.me.getMemberGroups())

      const announcements = await Promise.all(
        spItems.map(async (item): Promise<IAnnouncement> => {
          const targetUserIds = item.OData__ModernAudienceTargetUserFieldId

          let hasAccess = true
          if (props.targetAudience) {
            if (targetUserIds) {
              const groupIds = await Promise.all(
                targetUserIds.map(async (id) => {
                  const { LoginName } = await sp.web.getUserById(id)()
                  return LoginName.split('|').pop()
                })
              )
              hasAccess = groupIds.some((id) => memberOfGroups.includes(id))
            }
          }

          let severity: MessageBarIntent = 'info'

          switch (item.PzlSeverity) {
            case strings.Severity.Success:
              severity = 'success'
              break
            case strings.Severity.Warning:
              severity = 'warning'
              break
            case strings.Severity.Error:
              severity = 'error'
              break
            default:
              severity = 'info'
              break
          }

          return {
            id: String(item.ID),
            title: item.Title,
            severity,
            content: item.PzlContent,
            hasAccess,
            pinned: item.PzlPinned === true,
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

  /**
   * Fetches schema, permissions and announcements. Schema and permissions
   * never throw (they fall back to the template fields / `false`).
   */
  const fetchData = async () => {
    const [fields, canCreate] = await Promise.all([service.getFields(), service.canCreateItems()])
    setState({ fields, canCreate })
    try {
      const announcements = await getAnnouncements(fields)
      setState({ announcements, loading: false })
    } catch (error) {
      setState({
        error: error instanceof Error ? error : new Error(String(error)),
        loading: false
      })
    }
  }

  useEffect(() => {
    fetchData()
  }, [])

  return { refresh: fetchData }
}
