import { useEffect, useRef, useState } from 'react'
import { useAnnouncementState } from './useAnnouncementState'
import { IAnnouncementProps } from './types'
import { useId } from '@fluentui/react-components'
import { useAnnouncementDataFetch } from './useAnnouncementDataFetch'
import { getFocusedAnnouncementId } from '../util/getFocusedAnnouncementId'
import { IAnnouncement } from './types'
import { Status, statusChoiceValue } from '../model'
import { STATUS_FIELD } from '../constants'

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

  /**
   * Quick actions (status, edit) are shown to users who may edit items, when
   * the list has the status column
   */
  const statusField = state.fields.find((field) => field.internalName === STATUS_FIELD)
  const showQuickActions = state.canEdit && !!statusField

  /**
   * Sets the status of an announcement (quick action), closes its popover
   * and fetches the announcements again
   *
   * @param announcement Announcement
   * @param status New status
   */
  const setStatus = async (announcement: IAnnouncement, status: Status): Promise<void> => {
    await props.service.updateItem(announcement.id, {
      [statusField.entityPropertyName]: statusChoiceValue(status, statusField.choices)
    })
    setOpenAnnouncementId(null)
    await refresh().catch(() => {
      // Ignore, the list is shown as it was
    })
  }

  const getEditFormUrl = (announcement: IAnnouncement) =>
    props.service.getEditFormUrl(announcement.id)

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
    showQuickActions,
    setStatus,
    getEditFormUrl,
    createDialogOpen,
    setCreateDialogOpen
  }
}
