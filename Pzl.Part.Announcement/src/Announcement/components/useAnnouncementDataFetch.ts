import { useEffect } from 'react'
import { IAnnouncementProps, IAnnouncementState } from './types'
import { fetchAnnouncements } from '../services/fetchAnnouncements'

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

  /**
   * Fetches schema, permissions and announcements. Schema and permissions
   * never throw (they fall back to the template fields / `false`).
   */
  const fetchData = async () => {
    const [fields, canCreate] = await Promise.all([service.getFields(), service.canCreateItems()])
    setState({ fields, canCreate })
    try {
      const announcements = await fetchAnnouncements(
        service,
        props.context,
        fields,
        props.targetAudience
      )
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
