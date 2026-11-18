import { useEffect, useRef, useState } from 'react'
import { useAnnouncementState } from './useAnnouncementState'
import { IAnnouncementProps } from './types'
import { useId } from '@fluentui/react-components'
import { useAnnouncementDataFetch } from './useAnnouncementDataFetch'
import { getFocusedAnnouncementId } from '../util/getFocusedAnnouncementId'
import { IAnnouncementEditor } from './AnnouncementDialogs'
import { onAnnouncementChanged } from '../util/announcementChanged'

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
   * Identifies this web part when it tells the banner about a change
   */
  const partId = useId('announcement-web-part')

  // Read the list again when the banner on the page has changed an announcement
  useEffect(
    () =>
      onAnnouncementChanged(partId, () => {
        refresh().catch(() => {
          // The list is shown as it was
        })
      }),
    []
  )

  /**
   * ID of the announcement shown in the details dialog
   */
  const [openAnnouncementId, $setOpenAnnouncementId] = useState<string>(null)

  /**
   * Opens the details dialog for an announcement, or closes it (`null`)
   */
  const setOpenAnnouncementId = (id: string | null) => $setOpenAnnouncementId(id)

  const openAnnouncement = state.announcements.find((a) => a.id === openAnnouncementId)

  /**
   * Whether the create dialog is open
   */
  const [createDialogOpen, setCreateDialogOpen] = useState(false)

  /**
   * Refs to the announcement elements, used to scroll the focused one into view
   */
  const announcementRefs = useRef<Record<string, HTMLElement>>({})

  const setAnnouncementRef = (id: string) => (element: HTMLElement) => {
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
   * The editor functions of the dialogs, for users who may edit items
   */
  const editor: IAnnouncementEditor | undefined = state.canEdit
    ? {
        service: props.service,
        fields: state.fields,
        hiddenFormFields: props.formFields,
        alwaysNotified: state.alwaysNotified
      }
    : undefined

  return {
    state,
    setState,
    refresh,
    partId,
    fluentProviderId,
    toasterId,
    openAnnouncement,
    setOpenAnnouncementId,
    setAnnouncementRef,
    popoverFields,
    showCreateButton,
    editor,
    createDialogOpen,
    setCreateDialogOpen
  }
}
