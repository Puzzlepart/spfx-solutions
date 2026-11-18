import React, { FC, useState } from 'react'
import {
  Accordion,
  AccordionHeader,
  AccordionItem,
  AccordionPanel,
  Button,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  Divider,
  MessageBar,
  MessageBarBody,
  Spinner
} from '@fluentui/react-components'
import { EditRegular } from '@fluentui/react-icons'
import strings from 'AnnouncementStrings'
import { Status } from '../../model'
import { IAnnouncement, IListField } from '../types'
import { NotificationSettings } from '../NotificationSettings'
import styles from './EditorSection.module.scss'

export interface IEditorSectionProps {
  announcement: IAnnouncement
  /**
   * Whether the status buttons are shown (the list has the status column)
   */
  showStatusActions: boolean
  /**
   * Changes the status of the announcement. Rejects on failure.
   */
  onSetStatus: (announcement: IAnnouncement, status: Status) => Promise<void>
  /**
   * What «Avslutt melding» will do to this announcement, or nothing when
   * there is nothing to end (the button is then hidden)
   */
  endEffect?: { setsResolved: boolean; setsEndDate: boolean }
  /**
   * Ends the announcement (see `endEffect`). Rejects on failure.
   */
  onEnd: (announcement: IAnnouncement) => Promise<void>
  /**
   * Opens the announcement in the form
   */
  onEdit: (announcement: IAnnouncement) => void
  /**
   * «Varsle i Teams» (`PzlNotifyTeams`), if the list has it
   */
  notifyTeamsField?: IListField
  /**
   * «Varsle på nytt» (`PzlNotifyAgain`), if the list has it
   */
  notifyAgainField?: IListField
  /**
   * The announcement's level is one of «Nivåer som alltid varsles»
   */
  alwaysNotified: boolean
  /**
   * Saves a changed notification switch. Rejects on failure.
   */
  onSetNotification: (
    announcement: IAnnouncement,
    field: IListField,
    value: boolean
  ) => Promise<void>
}

/**
 * «Handlinger og varsling»: a collapsed section at the bottom of the details dialog,
 * shown only to users who may edit the list. Holds the status buttons, «Avslutt
 * melding» (with a confirmation), «Rediger» and the Teams notification switches.
 */
export const EditorSection: FC<IEditorSectionProps> = (props) => {
  const { announcement } = props
  const [busy, setBusy] = useState<Status>(null)
  const [statusFailed, setStatusFailed] = useState(false)
  const [confirmEnd, setConfirmEnd] = useState(false)
  const [ending, setEnding] = useState(false)
  const [endFailed, setEndFailed] = useState(false)

  const end = () => {
    setEnding(true)
    setEndFailed(false)
    props
      .onEnd(announcement)
      .then(() => setConfirmEnd(false))
      .catch(() => setEndFailed(true))
      .finally(() => setEnding(false))
  }

  const setStatus = (status: Status) => {
    setBusy(status)
    setStatusFailed(false)
    props
      .onSetStatus(announcement, status)
      .catch(() => setStatusFailed(true))
      .finally(() => setBusy(null))
  }

  const spinner = (status: Status) => (busy === status ? <Spinner size='tiny' /> : undefined)

  return (
    <div className={styles.editorSection}>
      <Divider />
      <Accordion collapsible>
        <AccordionItem value='editor'>
          <AccordionHeader expandIconPosition='end'>{strings.Editor.SectionTitle}</AccordionHeader>
          <AccordionPanel>
            <div className={styles.panel}>
              <div className={styles.actions}>
                {props.showStatusActions && announcement.status !== 'resolved' && (
                  <Button
                    appearance='primary'
                    disabled={!!busy}
                    icon={spinner('resolved')}
                    onClick={() => setStatus('resolved')}>
                    {strings.QuickActions.MarkResolved}
                  </Button>
                )}
                {props.showStatusActions && announcement.status === 'planned' && (
                  <Button
                    appearance='secondary'
                    disabled={!!busy}
                    icon={spinner('ongoing')}
                    onClick={() => setStatus('ongoing')}>
                    {strings.QuickActions.MarkOngoing}
                  </Button>
                )}
                {props.endEffect && (
                  <Button
                    appearance='secondary'
                    disabled={!!busy}
                    onClick={() => {
                      setEndFailed(false)
                      setConfirmEnd(true)
                    }}>
                    {strings.Editor.EndButtonText}
                  </Button>
                )}
                <Button
                  appearance='secondary'
                  disabled={!!busy}
                  icon={<EditRegular />}
                  onClick={() => props.onEdit(announcement)}>
                  {strings.QuickActions.Edit}
                </Button>
              </div>
              {statusFailed && (
                <MessageBar intent='error'>
                  <MessageBarBody>{strings.QuickActions.StatusChangeErrorText}</MessageBarBody>
                </MessageBar>
              )}
              {(props.notifyTeamsField || props.notifyAgainField || announcement.notifiedAt) && (
                <NotificationSettings
                  announcement={announcement}
                  notifyTeamsField={props.notifyTeamsField}
                  notifyAgainField={props.notifyAgainField}
                  alwaysNotified={props.alwaysNotified}
                  onChange={props.onSetNotification}
                />
              )}
            </div>
          </AccordionPanel>
        </AccordionItem>
      </Accordion>
      <Dialog
        modalType='alert'
        open={confirmEnd}
        onOpenChange={(_, data) => !ending && setConfirmEnd(data.open)}>
        <DialogSurface>
          <DialogBody>
            <DialogTitle>{strings.Editor.EndDialogTitle}</DialogTitle>
            <DialogContent>
              <p className={styles.confirmText}>
                {[
                  props.endEffect?.setsResolved && strings.Editor.EndSetsResolvedText,
                  props.endEffect?.setsEndDate && strings.Editor.EndSetsEndDateText
                ]
                  .filter(Boolean)
                  .join(' ')}
              </p>
              {endFailed && (
                <MessageBar intent='error'>
                  <MessageBarBody>{strings.Editor.EndErrorText}</MessageBarBody>
                </MessageBar>
              )}
            </DialogContent>
            <DialogActions>
              <Button appearance='secondary' disabled={ending} onClick={() => setConfirmEnd(false)}>
                {strings.CancelButtonText}
              </Button>
              <Button
                appearance='primary'
                disabled={ending}
                icon={ending ? <Spinner size='tiny' /> : undefined}
                onClick={end}>
                {strings.Editor.EndConfirmButtonText}
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
    </div>
  )
}
