/* eslint-disable prefer-spread */
import { useEffect, useRef, useState } from 'react'
import { useAnnouncementState } from './useAnnouncementState'
import { IAnnouncementProps } from './types'
import { useId } from '@fluentui/react-components'
import { useAnnouncementDataFetch } from './useAnnouncementDataFetch'
import { getFocusedAnnouncementId } from '../util/getFocusedAnnouncementId'

/**
 * Delay (ms) between scrolling the focused announcement into view and opening
 * its popover. The popover uses `closeOnScroll`, so it must not be open while
 * the scroll event from `scrollIntoView` fires.
 */
const OPEN_FOCUSED_DELAY = 250

/**
 * Component logic hook for `Announcement`. This hook is responsible for
 * fetching data, sorting, filtering and other logic.
 *
 * @param props Props
 */
export const useAnnouncement = (props: IAnnouncementProps) => {
  const { state, setState } = useAnnouncementState()
  useAnnouncementDataFetch(props, setState)

  const fluentProviderId = useId('fp-announcement')

  /**
   * ID of the announcement whose popover is currently open (controlled popovers)
   */
  const [openAnnouncementId, setOpenAnnouncementId] = useState<string>(null)

  /**
   * Refs to the announcement elements, used to scroll the focused one into view
   */
  const announcementRefs = useRef<Record<string, HTMLDivElement>>({})

  const setAnnouncementRef = (id: string) => (element: HTMLDivElement) => {
    announcementRefs.current[id] = element
  }

  // Resolve the announcement to focus from the Teams deep link or the URL
  useEffect(() => {
    getFocusedAnnouncementId(props.context)
      .then((focusedAnnouncementId) => {
        if (focusedAnnouncementId) setState({ focusedAnnouncementId })
      })
      .catch(() => {
        // Ignore, no announcement to focus
      })
  }, [])

  // Open the focused announcement once both the data and the focus ID are ready
  useEffect(() => {
    if (state.loading || !state.focusedAnnouncementId) return
    const match = state.announcements.find((a) => a.id === state.focusedAnnouncementId)
    if (!match) return
    const element = announcementRefs.current[match.id]
    if (element?.scrollIntoView) element.scrollIntoView({ block: 'center' })
    const timer = setTimeout(() => setOpenAnnouncementId(match.id), OPEN_FOCUSED_DELAY)
    return () => clearTimeout(timer)
  }, [state.loading, state.focusedAnnouncementId])

  return {
    state,
    setState,
    fluentProviderId,
    openAnnouncementId,
    setOpenAnnouncementId,
    setAnnouncementRef
  }
}
