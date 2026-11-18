import React, { FC } from 'react'
import { MessageBarIntent } from '@fluentui/react-components'
import {
  CalendarClockRegular,
  CheckmarkCircleFilled,
  DiamondDismissFilled,
  InfoFilled,
  WarningFilled
} from '@fluentui/react-icons'
import strings from 'AnnouncementStrings'
import { IAnnouncement } from '../types'
import { getIntent } from '../intent'
import { SEVERITY_FIELD } from '../../constants'
import styles from './AnnouncementDialog.module.scss'

/**
 * The icons and colours Fluent's `MessageBar` uses for each intent, so the
 * dialog shows the same icon as the announcement in the web part and the banner
 */
const ICONS: Record<MessageBarIntent, typeof InfoFilled> = {
  info: InfoFilled,
  warning: WarningFilled,
  error: DiamondDismissFilled,
  success: CheckmarkCircleFilled
}

export interface IIntentIconProps {
  announcement: IAnnouncement
}

/**
 * Icon for the level of an announcement, beside the dialog's title: the same
 * icon and colour as in the list. A resolved announcement gets the green check
 * mark and a planned one the calendar, as in the list.
 */
export const IntentIcon: FC<IIntentIconProps> = ({ announcement }) => {
  const intent = getIntent(announcement)
  const Icon = announcement.status === 'planned' ? CalendarClockRegular : ICONS[intent]
  // What the icon says: the status when it decides the icon, otherwise the level
  const label =
    announcement.status === 'resolved'
      ? strings.Status.Resolved
      : (announcement.values[SEVERITY_FIELD] ?? '')
  return (
    <Icon
      className={`${styles.intentIcon} ${styles[intent]}`}
      title={label || undefined}
      aria-label={label || undefined}
      aria-hidden={label ? undefined : true}
    />
  )
}
