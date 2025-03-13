import { useEffect } from 'react'
import { IAnnouncementProps, IAnnouncementState } from './types'
import { spfi, SPFx } from '@pnp/sp'
import '@pnp/sp/webs'
import '@pnp/sp/lists'
import '@pnp/sp/items'

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

      const announcementList = sp.web.lists.getByTitle('Announcement')
      const spItems = await announcementList.items.select('Title', 'Description')()

      console.log(spItems)

      return spItems.map((item) => {
        return {
          title: item.Title,
          description: item.Description
        }
      })
    } catch (error) {
      throw new Error('Kunne ikke hente meldinger...')
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
