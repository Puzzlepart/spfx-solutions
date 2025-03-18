import { useEffect } from 'react'
import { IAnnouncement, IAnnouncementProps, IAnnouncementState } from './types'
import { spfi, SPFx } from '@pnp/sp'
import '@pnp/sp/webs'
import '@pnp/sp/lists'
import '@pnp/sp/items'
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
      const sp = spfi().using(SPFx(props.context))
      const now = new Date()
      const dateFilter = `(PzlStartDate le datetime'${now.toISOString()}') and (PzlEndDate ge datetime'${now.toISOString()})'`

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
          'PzlResponsible/EMail'
        )
        // .filter(dateFilter)
        .expand('PzlResponsible')()

      console.log(spItems)

      return spItems.map((item): IAnnouncement => {
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
          responsible: { name: item.PzlResponsible?.Title, email: item.PzlResponsible?.EMail }
        }
      })
    } catch (error) {
      throw new Error('Kunne ikke hente driftsmeldinger...')
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
