import React, { FC, useState } from 'react'
import { Divider, Text } from '@fluentui/react-components'
import { format } from '@fluentui/react'
import strings from 'AnnouncementStrings'
import { IAnnouncement, IListField } from '../types'
import { FieldInput } from '../CreateAnnouncementDialog/FieldInput'
import { SEVERITY_FIELD } from '../../constants'
import styles from './NotificationSettings.module.scss'

export interface INotificationSettingsProps {
  announcement: IAnnouncement
  /**
   * «Varsle i Teams» (`PzlNotifyTeams`), if the list has it
   */
  notifyTeamsField?: IListField
  /**
   * «Varsle på nytt» (`PzlNotifyAgain`), if the list has it
   */
  notifyAgainField?: IListField
  /**
   * The announcement's level is one of «Nivåer som alltid varsles» in the
   * notification configuration: «Varsle i Teams» is shown as selected and
   * cannot be turned off, because the service notifies anyway.
   */
  alwaysNotified: boolean
  /**
   * Saves a changed switch on the announcement. Rejects on failure.
   */
  onChange: (announcement: IAnnouncement, field: IListField, value: boolean) => Promise<void>
}

/**
 * The switches of the notification add-on in the details dialog, for users
 * who may edit the list: «Varsle i Teams» and «Varsle på nytt». A change is
 * saved at once; the notification service picks it up on its next run.
 */
export const NotificationSettings: FC<INotificationSettingsProps> = ({
  announcement,
  notifyTeamsField,
  notifyAgainField,
  alwaysNotified,
  onChange
}) => {
  /**
   * Internal name of the field being saved, with the value it is changed to
   */
  const [saving, setSaving] = useState<{ internalName: string; value: boolean }>(null)
  const [failed, setFailed] = useState<string>(null)

  const change = (field: IListField, value: boolean) => {
    setSaving({ internalName: field.internalName, value })
    setFailed(null)
    onChange(announcement, field, value)
      .catch(() => setFailed(field.internalName))
      .finally(() => setSaving(null))
  }

  const renderSwitch = (field: IListField | undefined, locked: boolean) => {
    if (!field) return null
    const stored = announcement.values[field.entityPropertyName] === true
    const value = saving?.internalName === field.internalName ? saving.value : stored
    return (
      <FieldInput
        field={field}
        value={locked || value}
        hint={
          locked
            ? format(strings.Notification.AlwaysNotifiedHint, announcement.values[SEVERITY_FIELD])
            : undefined
        }
        error={failed === field.internalName ? strings.Notification.ChangeErrorText : undefined}
        disabled={locked || !!saving}
        onChange={(checked) => change(field, checked === true)}
      />
    )
  }

  return (
    <div className={styles.notificationSettings}>
      <Divider />
      <Text weight='semibold' block>
        {strings.Notification.SectionTitle}
      </Text>
      {renderSwitch(notifyTeamsField, alwaysNotified)}
      {renderSwitch(notifyAgainField, false)}
    </div>
  )
}
