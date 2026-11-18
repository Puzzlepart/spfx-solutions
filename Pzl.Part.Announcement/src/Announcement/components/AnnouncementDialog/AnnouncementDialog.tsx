import React, { FC, ReactNode } from 'react'
import {
  Button,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  DialogTrigger
} from '@fluentui/react-components'
import { Dismiss24Regular } from '@fluentui/react-icons'
import strings from 'AnnouncementStrings'
import { IAnnouncement, IListField } from '../types'
import { AnnouncementDetails } from '../AnnouncementDetails'
import { IntentIcon } from './IntentIcon'
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
   * Shown below the details: «Handlinger og varsling», for users who may edit the list
   */
  children?: ReactNode
}

/**
 * The details of an announcement in a modal dialog at the centre of the page.
 * Used through `AnnouncementDialogs` by the web part and the banner.
 */
export const AnnouncementDialog: FC<IAnnouncementDialogProps> = ({
  announcement,
  fields,
  onClose,
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
            <span className={styles.title}>
              <IntentIcon announcement={announcement} />
              <span>{announcement.title}</span>
            </span>
          </DialogTitle>
          <DialogContent className={styles.content}>
            <AnnouncementDetails announcement={announcement} fields={fields} />
            {children}
          </DialogContent>
          <DialogActions>
            <DialogTrigger action='close' disableButtonEnhancement>
              <Button appearance='secondary'>{strings.CloseButtonText}</Button>
            </DialogTrigger>
          </DialogActions>
        </DialogBody>
      )}
    </DialogSurface>
  </Dialog>
)
