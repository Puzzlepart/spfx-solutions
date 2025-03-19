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
 * Component data fetch hook for `Announcement`. This hook is responsible for
 * fetching data and setting state.
 *
 * @param props Props
 * @param refetch Timestamp for refetch. Changes to this variable refetches the data in `useEffect`
 * @param setState Set state callback
 */
export function useAnnouncementDataFetch(
  props: IAnnouncementProps,
  setState: (newState: Partial<IAnnouncementState>) => void
) {
  const getAnnouncements = async (): Promise<any[]> => {
    try {
      const graph = graphfi().using(graphSPFx(props.context))
      const sp = spfi().using(SPFx(props.context))
      const now = new Date()
      const announcementList = sp.web.lists.getByTitle(strings.AnnouncementsListName)
      const spItems = await announcementList.items
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
          props.targetAudience ? 'OData__ModernAudienceTargetUserFieldId' : '*'
        )
        .filter(
          `PzlStartDate le datetime'${now.toISOString()}' and PzlEndDate ge datetime'${now.toISOString()}'`
        )
        .expand('PzlResponsible')()

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
            id: item.ID,
            title: item.Title,
            severity,
            content: item.PzlContent,
            startDate: item.PzlStartDate,
            endDate: item.PzlEndDate,
            affectedSystems: item.PzlAffectedSystems,
            consequence: item.PzlConsequences,
            responsible: { name: item.PzlResponsible?.Title, email: item.PzlResponsible?.EMail },
            hasAccess
          }
        })
      )
      return announcements.filter((announcement) => announcement.hasAccess)
    } catch (error) {
      throw new Error('Kunne ikke hente driftsmeldinger...')
    }
  }

  useEffect(() => {
    const fetchData = async () => {
      try {
        const announcements = await getAnnouncements()
        console.log(announcements)
        setState({ announcements, loading: false })
      } catch (error) {
        setState({ error, loading: false })
      }
    }

    fetchData()
  }, [])
}
