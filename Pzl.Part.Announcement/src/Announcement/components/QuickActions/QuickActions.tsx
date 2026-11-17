import React, { FC, useState } from 'react'
import { Button, Spinner } from '@fluentui/react-components'
import { EditRegular } from '@fluentui/react-icons'
import strings from 'AnnouncementStrings'
import { Status } from '../../model'
import { IAnnouncement } from '../types'

export interface IStatusActionsProps {
  announcement: IAnnouncement
  /**
   * Changes the status of the announcement. Rejects on failure.
   */
  onSetStatus: (announcement: IAnnouncement, status: Status) => Promise<void>
}

/**
 * Status buttons in the details dialog for users who may edit the list: move
 * a planned announcement to ongoing, or mark it as resolved. Placed at the end
 * of the dialog's actions, with the primary button («Marker som løst») last.
 */
export const StatusActions: FC<IStatusActionsProps> = ({ announcement, onSetStatus }) => {
  const [busy, setBusy] = useState<Status>(null)

  const setStatus = (status: Status) => {
    setBusy(status)
    onSetStatus(announcement, status)
      .catch(() => {
        // The caller shows the error
      })
      .finally(() => setBusy(null))
  }

  const spinner = (status: Status) => (busy === status ? <Spinner size='tiny' /> : undefined)

  return (
    <>
      {announcement.status === 'planned' && (
        <Button
          appearance='secondary'
          disabled={!!busy}
          icon={spinner('ongoing')}
          onClick={() => setStatus('ongoing')}>
          {strings.QuickActions.MarkOngoing}
        </Button>
      )}
      {announcement.status !== 'resolved' && (
        <Button
          appearance='primary'
          disabled={!!busy}
          icon={spinner('resolved')}
          onClick={() => setStatus('resolved')}>
          {strings.QuickActions.MarkResolved}
        </Button>
      )}
    </>
  )
}

export interface IEditActionProps {
  announcement: IAnnouncement
  /**
   * Returns the URL of the list's edit form for the announcement
   */
  getEditFormUrl: (announcement: IAnnouncement) => Promise<string | null>
}

/**
 * «Rediger» in the details dialog: opens the list's edit form, which is used
 * for anything but the status and the notification. Placed at the start of
 * the dialog's actions.
 */
export const EditAction: FC<IEditActionProps> = ({ announcement, getEditFormUrl }) => {
  const openEditForm = () => {
    getEditFormUrl(announcement)
      .then((url) => {
        if (url) window.location.href = url
      })
      .catch(() => {
        // No edit form URL, nothing to open
      })
  }

  return (
    <Button appearance='subtle' icon={<EditRegular />} onClick={openEditForm}>
      {strings.QuickActions.Edit}
    </Button>
  )
}
