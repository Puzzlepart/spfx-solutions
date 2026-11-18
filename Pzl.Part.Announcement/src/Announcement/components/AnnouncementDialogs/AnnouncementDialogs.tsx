import React, { FC, useState } from 'react'
import { Toast, Toaster, ToastTitle, useId, useToastController } from '@fluentui/react-components'
import { format } from '@fluentui/react'
import strings from 'AnnouncementStrings'
import { IAnnouncement, IListField } from '../types'
import { AnnouncementDialog } from '../AnnouncementDialog'
import { EditorSection } from '../EditorSection'
import { CreateAnnouncementDialog } from '../CreateAnnouncementDialog'
import { statusText } from '../StatusBadge'
import { AnnouncementListService } from '../../services/AnnouncementListService'
import { isAlwaysNotified, Severity, Status, statusChoiceValue } from '../../model'
import {
  NOTIFY_AGAIN_FIELD,
  NOTIFY_TEAMS_FIELD,
  SEVERITY_FIELD,
  STATUS_FIELD
} from '../../constants'

/**
 * What the editor functions need. Given only for users who may edit the list.
 */
export interface IAnnouncementEditor {
  service: AnnouncementListService
  /**
   * All fields of the list (the form shows the editable ones)
   */
  fields: IListField[]
  /**
   * Fields hidden from the form in the web part properties (`false` = hidden)
   */
  hiddenFormFields?: Record<string, boolean>
  /**
   * Levels the notification add-on always notifies about
   */
  alwaysNotified: Severity[]
}

export interface IAnnouncementDialogsProps {
  /**
   * The announcement shown in the details dialog. The dialog is open while this is set.
   */
  announcement: IAnnouncement | undefined
  /**
   * Fields shown in the details, in order
   */
  fields: IListField[]
  onClose: () => void
  /**
   * Set for users who may edit the list: adds «For redaktører» and the edit form
   */
  editor?: IAnnouncementEditor
  /**
   * Called after an announcement was changed, to read the announcements again
   */
  onChanged: () => Promise<void>
}

/**
 * Everything that opens from a click on an announcement, the same in the web
 * part and the banner: the details dialog, and for editors the section «For
 * redaktører» (status, edit, Teams notification), the edit form and the
 * confirmations.
 */
export const AnnouncementDialogs: FC<IAnnouncementDialogsProps> = ({
  announcement,
  fields,
  onClose,
  editor,
  onChanged
}) => {
  const toasterId = useId('toaster-announcement-dialogs')
  const { dispatchToast } = useToastController(toasterId)
  /**
   * The announcement in the edit form. Kept after the form closes, so the form
   * does not change while it fades out.
   */
  const [editAnnouncement, setEditAnnouncement] = useState<IAnnouncement>(null)
  const [formOpen, setFormOpen] = useState(false)

  const field = (internalName: string) =>
    editor?.fields.find((f) => f.internalName === internalName)
  const statusField = field(STATUS_FIELD)

  const toast = (text: string) =>
    dispatchToast(
      <Toast>
        <ToastTitle>{text}</ToastTitle>
      </Toast>,
      { intent: 'success' }
    )

  const refresh = () =>
    onChanged().catch(() => {
      // The announcements are shown as they were
    })

  /**
   * Sets the status, closes the dialog and reads the announcements again
   */
  const setStatus = async (item: IAnnouncement, status: Status): Promise<void> => {
    await editor.service.updateItem(item.id, {
      [statusField.entityPropertyName]: statusChoiceValue(status, statusField.choices)
    })
    onClose()
    toast(format(strings.QuickActions.StatusChangedText, statusText(status).toLowerCase()))
    await refresh()
  }

  /**
   * Saves «Varsle i Teams» or «Varsle på nytt». The dialog stays open.
   */
  const setNotification = async (
    item: IAnnouncement,
    notifyField: IListField,
    value: boolean
  ): Promise<void> => {
    await editor.service.updateItem(item.id, { [notifyField.entityPropertyName]: value })
    await refresh()
  }

  /**
   * Closes the details and opens the announcement in the form
   */
  const edit = (item: IAnnouncement) => {
    onClose()
    setEditAnnouncement(item)
    setFormOpen(true)
  }

  const onSaved = () => {
    setFormOpen(false)
    toast(strings.EditSuccessText)
    refresh().catch(() => undefined)
  }

  return (
    <>
      <AnnouncementDialog announcement={announcement} fields={fields} onClose={onClose}>
        {editor && announcement && (
          <EditorSection
            announcement={announcement}
            showStatusActions={!!statusField}
            onSetStatus={setStatus}
            onEdit={edit}
            notifyTeamsField={field(NOTIFY_TEAMS_FIELD)}
            notifyAgainField={field(NOTIFY_AGAIN_FIELD)}
            alwaysNotified={isAlwaysNotified(
              announcement.values[SEVERITY_FIELD],
              editor.alwaysNotified
            )}
            onSetNotification={setNotification}
          />
        )}
      </AnnouncementDialog>
      {editor && editAnnouncement && (
        <CreateAnnouncementDialog
          open={formOpen}
          fields={editor.fields}
          hiddenFormFields={editor.hiddenFormFields}
          alwaysNotified={editor.alwaysNotified}
          service={editor.service}
          announcement={editAnnouncement}
          onOpenChange={setFormOpen}
          onSaved={onSaved}
        />
      )}
      <Toaster toasterId={toasterId} position='bottom-end' />
    </>
  )
}
