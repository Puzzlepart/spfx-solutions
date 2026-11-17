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
   * Buttons for users who may edit the list (status, edit), shown to the left
   * of the close button
   */
  actions?: ReactNode
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
  actions,
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
          {actions && <DialogActions position='start'>{actions}</DialogActions>}
          <DialogActions position='end'>
            <DialogTrigger action='close'>
              <Button appearance='secondary'>{strings.CloseButtonText}</Button>
            </DialogTrigger>
          </DialogActions>
        </DialogBody>
      )}
    </DialogSurface>
  </Dialog>
)
