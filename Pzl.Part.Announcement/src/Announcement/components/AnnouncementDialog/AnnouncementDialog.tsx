import React, { FC, ReactNode } from 'react'
import {
  Button,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  DialogTrigger,
  MessageBar,
  MessageBarBody
} from '@fluentui/react-components'
import { Dismiss24Regular } from '@fluentui/react-icons'
import strings from 'AnnouncementStrings'
import { IAnnouncement, IListField } from '../types'
import { AnnouncementDetails } from '../AnnouncementDetails'
import styles from './AnnouncementDialog.module.scss'

export interface IAnnouncementDialogProps {
  /**
   * The announcement to show. The dialog is open while this is set.
   */
  announcement: IAnnouncement | undefined
  /**
   * Fields to show, in order
   */
  fields: IListField[]
  onClose: () => void
  /**
   * Secondary action at the start (left) of the actions, e.g. «Rediger»
   */
  startActions?: ReactNode
  /**
   * Actions after the close button, at the end (right), e.g. the status
   * buttons. The primary button comes last, as in other Fluent dialogs.
   */
  endActions?: ReactNode
  /**
   * Error from an action, shown at the bottom of the content
   */
  error?: string
  /**
   * Shown below the details, e.g. the notification settings for editors
   */
  children?: ReactNode
}

/**
 * The details of an announcement in a modal dialog at the centre of the page.
 * Used by the web part (with the editor functions) and the banner.
 */
export const AnnouncementDialog: FC<IAnnouncementDialogProps> = ({
  announcement,
  fields,
  onClose,
  startActions,
  endActions,
  error,
  children
}) => (
  <Dialog open={!!announcement} onOpenChange={(_, data) => !data.open && onClose()}>
    <DialogSurface aria-describedby={undefined}>
      {announcement && (
        <DialogBody>
          <DialogTitle
            action={
              <DialogTrigger action='close'>
                <Button
                  appearance='subtle'
                  icon={<Dismiss24Regular />}
                  aria-label={strings.Aria.CloseDialog}
                />
              </DialogTrigger>
            }>
            {announcement.title}
          </DialogTitle>
          <DialogContent className={styles.content}>
            <AnnouncementDetails announcement={announcement} fields={fields} />
            {children}
            {error && (
              <MessageBar intent='error'>
                <MessageBarBody>{error}</MessageBarBody>
              </MessageBar>
            )}
          </DialogContent>
          {startActions && (
            <DialogActions position='start' className={styles.actions}>
              {startActions}
            </DialogActions>
          )}
          <DialogActions position='end' className={styles.actions}>
            <DialogTrigger action='close' disableButtonEnhancement>
              <Button appearance='secondary'>{strings.CloseButtonText}</Button>
            </DialogTrigger>
            {endActions}
          </DialogActions>
        </DialogBody>
      )}
    </DialogSurface>
  </Dialog>
)
