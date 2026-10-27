import { useEffect, useRef, useState } from 'react'
import { useAnnouncementState } from './useAnnouncementState'
import { IAnnouncementProps } from './types'
import { useId } from '@fluentui/react-components'
import { useAnnouncementDataFetch } from './useAnnouncementDataFetch'
import { getFocusedAnnouncementId, getPageRedirectUrl } from '../util/getFocusedAnnouncementId'

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
  const { refresh } = useAnnouncementDataFetch(props, setState)

  const fluentProviderId = useId('fp-announcement')
  const toasterId = useId('toaster-announcement')

  /**
   * ID of the announcement whose popover is currently open (controlled popovers)
   */
  const [openAnnouncementId, setOpenAnnouncementId] = useState<string>(null)

  /**
   * Whether the create dialog is open
   */
  const [createDialogOpen, setCreateDialogOpen] = useState(false)

  /**
   * Toggles the popover for an announcement (keyboard activation of the trigger)
   */
  const toggleAnnouncement = (id: string) =>
    setOpenAnnouncementId((current) => (current === id ? null : id))

  /**
   * Refs to the announcement elements, used to scroll the focused one into view
   */
  const announcementRefs = useRef<Record<string, HTMLDivElement>>({})

  const setAnnouncementRef = (id: string) => (element: HTMLDivElement) => {
    announcementRefs.current[id] = element
  }

  // A news deep link (`page:<url>`) lands on the page hosting the web part;
  // go to the news page itself (same origin only, see getPageRedirectUrl).
  // The AnnouncementDeepLink application customizer does the same before the
  // page renders; this is the fallback for sites where it is not registered.
  useEffect(() => {
    const url = getPageRedirectUrl()
    if (url) window.location.replace(url)
  }, [])

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

  /**
   * Fields selected for the popover (in content type order)
   */
  const popoverFields = state.fields.filter(
    (field) => props.popoverFields?.[field.internalName] === true
  )

  /**
   * The create button is shown to users who may add items to the list,
   * unless hidden in the web part properties.
   */
  const showCreateButton = state.canCreate && !props.hideCreateButton

  return {
    state,
    setState,
    refresh,
    fluentProviderId,
    toasterId,
    openAnnouncementId,
    setOpenAnnouncementId,
    setAnnouncementRef,
    toggleAnnouncement,
    popoverFields,
    showCreateButton,
    createDialogOpen,
    setCreateDialogOpen
  }
}
