import React, { FC } from 'react'
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
  MessageBarBody,
  Spinner
} from '@fluentui/react-components'
import { Dismiss24Regular } from '@fluentui/react-icons'
import strings from 'AnnouncementStrings'
import styles from './CreateAnnouncementDialog.module.scss'
import { IAnnouncement, IListField } from '../types'
import { Severity } from '../../model'
import { AnnouncementListService } from '../../services/AnnouncementListService'
import { useCreateAnnouncementForm } from './useCreateAnnouncementForm'
import { FieldInput } from './FieldInput'

export interface ICreateAnnouncementDialogProps {
  open: boolean
  /**
   * All list fields; the form shows the editable ones
   */
  fields: IListField[]
  /**
   * Fields hidden from the form in the web part properties (`false` = hidden)
   */
  hiddenFormFields?: Record<string, boolean>
  /**
   * Levels the notification add-on always notifies about: «Varsle i Teams» is
   * selected and locked for them
   */
  alwaysNotified?: Severity[]
  service: AnnouncementListService
  onOpenChange: (open: boolean) => void
  /**
   * The announcement to edit. Without it the form registers a new one.
   */
  announcement?: IAnnouncement
  /**
   * Called after the announcement has been created or updated
   */
  onSaved: () => void
}

/**
 * Dialog for registering a new announcement, or editing an existing one. The
 * form is built from the fields of the announcement list.
 */
export const CreateAnnouncementDialog: FC<ICreateAnnouncementDialogProps> = (props) => {
  const { formFields, values, lockedFields, errors, saving, saveError, setValue, submit } =
    useCreateAnnouncementForm(
      props.open,
      props.fields,
      props.hiddenFormFields,
      props.alwaysNotified ?? [],
      props.service,
      props.onSaved,
      props.announcement
    )

  const searchPeople = (query: string, allowGroups: boolean) =>
    props.service.searchPeople(query, allowGroups)

  return (
    <Dialog open={props.open} onOpenChange={(_, data) => props.onOpenChange(data.open)}>
      <DialogSurface aria-describedby={undefined}>
        <form
          onSubmit={(event) => {
            event.preventDefault()
            submit().catch(() => {
              // Errors are shown in the form
            })
          }}>
          <DialogBody>
            <DialogTitle
              action={
                <DialogTrigger action='close'>
                  <Button
                    appearance='subtle'
                    aria-label={strings.Aria.CloseDialog}
                    icon={<Dismiss24Regular />}
                    disabled={saving}
                  />
                </DialogTrigger>
              }>
              {props.announcement ? strings.EditDialogTitle : strings.CreateDialogTitle}
            </DialogTitle>
            <DialogContent className={styles.form}>
              {formFields.map((field) => (
                <FieldInput
                  key={field.internalName}
                  field={field}
                  value={values[field.internalName]}
                  error={errors[field.internalName]}
                  disabled={saving || field.internalName in lockedFields}
                  hint={lockedFields[field.internalName]}
                  onChange={(value) => setValue(field.internalName, value)}
                  searchPeople={searchPeople}
                />
              ))}
              {saveError && (
                <MessageBar intent='error'>
                  <MessageBarBody>{saveError}</MessageBarBody>
                </MessageBar>
              )}
            </DialogContent>
            <DialogActions>
              <DialogTrigger disableButtonEnhancement>
                <Button appearance='secondary' disabled={saving}>
                  {strings.CancelButtonText}
                </Button>
              </DialogTrigger>
              <Button
                type='submit'
                appearance='primary'
                disabled={saving}
                icon={saving ? <Spinner size='tiny' /> : undefined}>
                {strings.SaveButtonText}
              </Button>
            </DialogActions>
          </DialogBody>
        </form>
      </DialogSurface>
    </Dialog>
  )
}
