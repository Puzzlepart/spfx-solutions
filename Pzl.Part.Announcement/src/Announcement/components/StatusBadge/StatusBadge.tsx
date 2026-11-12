import React, { FC } from 'react'
import { Badge } from '@fluentui/react-components'
import { CalendarClock16Regular, Checkmark16Regular } from '@fluentui/react-icons'
import strings from 'AnnouncementStrings'
import { Status } from '../../model'

export interface IStatusBadgeProps {
  status: Status
  className?: string
}

/**
 * Label for planned and resolved announcements. Ongoing announcements get no
 * label: that is the normal case, and the colour already shows the level.
 */
export const StatusBadge: FC<IStatusBadgeProps> = ({ status, className }) => {
  switch (status) {
    case 'planned':
      return (
        <Badge
          className={className}
          appearance='tint'
          color='informative'
          size='small'
          icon={<CalendarClock16Regular />}>
          {strings.Status.Planned}
        </Badge>
      )
    case 'resolved':
      return (
        <Badge
          className={className}
          appearance='tint'
          color='success'
          size='small'
          icon={<Checkmark16Regular />}>
          {strings.Status.Resolved}
        </Badge>
      )
    default:
      return null
  }
}

/**
 * Display name of a status
 */
export function statusText(status: Status): string {
  switch (status) {
    case 'planned':
      return strings.Status.Planned
    case 'resolved':
      return strings.Status.Resolved
    default:
      return strings.Status.Ongoing
  }
}
