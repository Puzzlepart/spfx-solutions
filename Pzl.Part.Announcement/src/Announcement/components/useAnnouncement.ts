import { useEffect, useRef, useState } from 'react'
import strings from 'AnnouncementStrings'
import { useAnnouncementState } from './useAnnouncementState'
import { IAnnouncementProps } from './types'
import { useId } from '@fluentui/react-components'
import { useAnnouncementDataFetch } from './useAnnouncementDataFetch'
import { getFocusedAnnouncementId } from '../util/getFocusedAnnouncementId'
import { IAnnouncement, IListField } from './types'
import { isAlwaysNotified, Status, statusChoiceValue } from '../model'
import { NOTIFY_AGAIN_FIELD, NOTIFY_TEAMS_FIELD, SEVERITY_FIELD, STATUS_FIELD } from '../constants'

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
   * ID of the announcement shown in the details dialog
   */
  const [openAnnouncementId, $setOpenAnnouncementId] = useState<string>(null)

  /**
   * Error from an action in the details dialog (status, notification)
   */
  const [dialogError, setDialogError] = useState<string>(null)

  /**
   * Opens the details dialog for an announcement, or closes it (`null`)
   */
  const setOpenAnnouncementId = (id: string | null) => {
    setDialogError(null)
    $setOpenAnnouncementId(id)
  }

  const openAnnouncement = state.announcements.find((a) => a.id === openAnnouncementId)

  /**
   * Whether the create dialog is open
   */
  const [createDialogOpen, setCreateDialogOpen] = useState(false)

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
    setOpenAnnouncementId(match.id)
  }, [state.loading, state.focusedAnnouncementId])

  /**
   * Fields selected for the details dialog (in content type order)
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
   * The status buttons are shown to users who may edit items, when the list
   * has the status column. «Rediger» is shown to all who may edit items.
   */
  const statusField = state.fields.find((field) => field.internalName === STATUS_FIELD)
  const showQuickActions = state.canEdit && !!statusField

  /**
   * Sets the status of an announcement (quick action), closes the dialog
   * and fetches the announcements again
   *
   * @param announcement Announcement
   * @param status New status
   */
  const setStatus = async (announcement: IAnnouncement, status: Status): Promise<void> => {
    setDialogError(null)
    try {
      await props.service.updateItem(announcement.id, {
        [statusField.entityPropertyName]: statusChoiceValue(status, statusField.choices)
      })
    } catch (error) {
      setDialogError(strings.QuickActions.StatusChangeErrorText)
      throw error
    }
    setOpenAnnouncementId(null)
    await refresh().catch(() => {
      // Ignore, the list is shown as it was
    })
  }

  /**
   * The notification add-on's switches are shown to users who may edit items,
   * when the list has the add-on's columns
   */
  const notifyTeamsField = state.fields.find((field) => field.internalName === NOTIFY_TEAMS_FIELD)
  const notifyAgainField = state.fields.find((field) => field.internalName === NOTIFY_AGAIN_FIELD)
  const showNotificationSettings = state.canEdit && (!!notifyTeamsField || !!notifyAgainField)

  /**
   * Whether the level of the open announcement is always notified
   */
  const openAlwaysNotified = isAlwaysNotified(
    openAnnouncement?.values[SEVERITY_FIELD],
    state.alwaysNotified
  )

  /**
   * Saves «Varsle i Teams» or «Varsle på nytt» on an announcement and fetches
   * the announcements again. The dialog stays open.
   *
   * @param announcement Announcement
   * @param field The switch's field
   * @param value New value
   */
  const setNotification = async (
    announcement: IAnnouncement,
    field: IListField,
    value: boolean
  ): Promise<void> => {
    await props.service.updateItem(announcement.id, { [field.entityPropertyName]: value })
    await refresh().catch(() => {
      // Ignore, the list is shown as it was
    })
  }

  /**
   * The announcement shown in the form dialog: set when «Rediger» opened it,
   * empty when it registers a new announcement
   */
  const [editAnnouncement, setEditAnnouncement] = useState<IAnnouncement>(null)

  /**
   * Opens the form for a new announcement
   */
  const openCreateDialog = () => {
    setEditAnnouncement(null)
    setCreateDialogOpen(true)
  }

  /**
   * Closes the details dialog and opens the announcement in the form
   */
  const openEditDialog = (announcement: IAnnouncement) => {
    setOpenAnnouncementId(null)
    setEditAnnouncement(announcement)
    setCreateDialogOpen(true)
  }

  return {
    state,
    setState,
    refresh,
    fluentProviderId,
    toasterId,
    openAnnouncement,
    setOpenAnnouncementId,
    dialogError,
    setAnnouncementRef,
    popoverFields,
    showCreateButton,
    showQuickActions,
    setStatus,
    editAnnouncement,
    openCreateDialog,
    openEditDialog,
    showNotificationSettings,
    notifyTeamsField,
    notifyAgainField,
    openAlwaysNotified,
    setNotification,
    createDialogOpen,
    setCreateDialogOpen
  }
}
