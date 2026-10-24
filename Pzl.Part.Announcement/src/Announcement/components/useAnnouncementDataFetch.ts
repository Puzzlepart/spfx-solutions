import { useEffect } from 'react'
import { IAnnouncement, IAnnouncementProps, IAnnouncementState } from './types'
import { graphfi, SPFx as graphSPFx } from '@pnp/graph'
import '@pnp/graph/users'
import '@pnp/graph/groups'
import '@pnp/graph/members'
import { spfi, SPFx } from '@pnp/sp'
import '@pnp/sp/webs'
import '@pnp/sp/lists'
import '@pnp/sp/items'
import '@pnp/sp/site-groups/web'
import '@pnp/sp/webs'
import '@pnp/sp/site-users/web'
import strings from 'AnnouncementStrings'
import { MessageBarIntent } from '@fluentui/react-components'

/**
 * Fields added in v1.1.0 (Teams notification). Lists provisioned with an older
 * template do not have them, so the query is retried without them on failure.
 */
const NOTIFICATION_FIELDS = ['PzlPinned', 'PzlNotifiedAt']

/**
 * Component data fetch hook for `Announcement`. This hook is responsible for
 * fetching data and setting state.
 *
 * @param props Props
 * @param setState Set state callback
 */
export function useAnnouncementDataFetch(
  props: IAnnouncementProps,
  setState: (newState: Partial<IAnnouncementState>) => void
) {
  const sp = spfi().using(SPFx(props.context))

  /**
   * Fetches active items from the announcement list.
   *
   * @param includeNotificationFields Include the fields added in v1.1.0
   */
  const fetchItems = (includeNotificationFields: boolean): Promise<any[]> => {
    const now = new Date().toISOString()
    const announcementList = sp.web.lists.getByTitle(strings.AnnouncementsListName)
    return announcementList.items
      .select(
        'ID',
        'Title',
        'PzlSeverity',
        'PzlContent',
        'PzlStartDate',
        'PzlEndDate',
        'PzlAffectedSystems',
        'PzlConsequences',
        'PzlResponsible/Title',
        'PzlResponsible/EMail',
        ...(includeNotificationFields ? NOTIFICATION_FIELDS : []),
        props.targetAudience ? 'OData__ModernAudienceTargetUserFieldId' : '*'
      )
      .filter(`PzlStartDate le datetime'${now}' and PzlEndDate ge datetime'${now}'`)
      .expand('PzlResponsible')()
  }

  const getAnnouncements = async (): Promise<IAnnouncement[]> => {
    try {
      const graph = graphfi().using(graphSPFx(props.context))

      let spItems: any[]
      try {
        spItems = await fetchItems(true)
      } catch {
        // List not updated with the v1.1.0 template, fall back to the original fields
        spItems = await fetchItems(false)
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
            startDate: item.PzlStartDate,
            endDate: item.PzlEndDate,
            affectedSystems: item.PzlAffectedSystems,
            consequence: item.PzlConsequences,
            responsible: { name: item.PzlResponsible?.Title, email: item.PzlResponsible?.EMail },
            hasAccess,
            pinned: item.PzlPinned === true,
            notifiedAt: item.PzlNotifiedAt ?? null
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

  useEffect(() => {
    const fetchData = async () => {
      try {
        const announcements = await getAnnouncements()
        setState({ announcements, loading: false })
      } catch (error) {
        setState({ error, loading: false })
      }
    }

    fetchData()
  }, [])
}
