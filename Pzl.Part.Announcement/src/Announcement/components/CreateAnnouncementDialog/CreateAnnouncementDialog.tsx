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
import { IListField } from '../types'
import { AnnouncementListService } from '../../services/AnnouncementListService'
import { useCreateAnnouncementForm } from './useCreateAnnouncementForm'
import { FieldInput } from './FieldInput'

export interface ICreateAnnouncementDialogProps {
  open: boolean
  /**
   * All list fields; the form shows the editable ones
   */
  fields: IListField[]
  service: AnnouncementListService
  onOpenChange: (open: boolean) => void
  /**
   * Called after the announcement has been created
   */
  onCreated: () => void
}

/**
 * Dialog for registering a new announcement. The form is built from the
 * fields of the announcement list.
 */
export const CreateAnnouncementDialog: FC<ICreateAnnouncementDialogProps> = (props) => {
  const { formFields, values, errors, saving, saveError, setValue, submit } =
    useCreateAnnouncementForm(props.open, props.fields, props.service, props.onCreated)

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
              {strings.CreateDialogTitle}
            </DialogTitle>
            <DialogContent className={styles.form}>
              {formFields.map((field) => (
                <FieldInput
                  key={field.internalName}
                  field={field}
                  value={values[field.internalName]}
                  error={errors[field.internalName]}
                  disabled={saving}
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
