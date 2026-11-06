import React, { FC, useState } from 'react'
import { Button, Link, Spinner } from '@fluentui/react-components'
import { Checkmark16Regular, Edit16Regular, Play16Regular } from '@fluentui/react-icons'
import strings from 'AnnouncementStrings'
import { Status } from '../../model'
import { IAnnouncement } from '../types'
import styles from './QuickActions.module.scss'

export interface IQuickActionsProps {
  announcement: IAnnouncement
  /**
   * Changes the status of the announcement. Rejects on failure.
   */
  onSetStatus: (announcement: IAnnouncement, status: Status) => Promise<void>
  /**
   * Returns the URL of the list's edit form for the announcement
   */
  getEditFormUrl: (announcement: IAnnouncement) => Promise<string | null>
}

/**
 * Quick actions in the popover for users who may edit the list: move a
 * planned announcement to ongoing, mark it as resolved, or open the list's
 * edit form. The list's own form is used for anything else.
 */
export const QuickActions: FC<IQuickActionsProps> = ({
  announcement,
  onSetStatus,
  getEditFormUrl
}) => {
  const [busy, setBusy] = useState<Status>(null)

  const setStatus = (status: Status) => {
    setBusy(status)
    onSetStatus(announcement, status)
      .catch(() => {
        // The caller shows the error
      })
      .finally(() => setBusy(null))
  }

  const openEditForm = (event: React.MouseEvent) => {
    event.preventDefault()
    getEditFormUrl(announcement)
      .then((url) => {
        if (url) window.location.href = url
      })
      .catch(() => {
        // No edit form URL, nothing to open
      })
  }

  const icon = (status: Status, fallback: JSX.Element) =>
    busy === status ? <Spinner size='extra-tiny' /> : fallback

  return (
    <div className={styles.quickActions}>
      {announcement.status === 'planned' && (
        <Button
          size='small'
          appearance='secondary'
          disabled={!!busy}
          icon={icon('ongoing', <Play16Regular />)}
          onClick={() => setStatus('ongoing')}>
          {strings.QuickActions.MarkOngoing}
        </Button>
      )}
      {announcement.status !== 'resolved' && (
        <Button
          size='small'
          appearance='primary'
          disabled={!!busy}
          icon={icon('resolved', <Checkmark16Regular />)}
          onClick={() => setStatus('resolved')}>
          {strings.QuickActions.MarkResolved}
        </Button>
      )}
      <Link className={styles.edit} href='#' onClick={openEditForm}>
        <Edit16Regular />
        {strings.QuickActions.Edit}
      </Link>
    </div>
  )
}
