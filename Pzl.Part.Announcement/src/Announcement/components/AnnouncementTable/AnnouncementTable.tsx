import React, { FC } from 'react'
import {
  Table,
  TableBody,
  TableCell,
  TableHeader,
  TableHeaderCell,
  TableRow
} from '@fluentui/react-components'
import { Pin16Filled } from '@fluentui/react-icons'
import { format } from '@fluentui/react'
import strings from 'AnnouncementStrings'
import { IAnnouncement, IListField } from '../types'
import { IntentIcon } from '../AnnouncementDialog/IntentIcon'
import { statusText } from '../StatusBadge'
import { SEVERITY_FIELD } from '../../constants'
import { formatNumericDate } from '../../util/formatDate'
import styles from './AnnouncementTable.module.scss'

export interface IAnnouncementTableProps {
  announcements: IAnnouncement[]
  /**
   * The list fields, to know whether the dates have a time of day
   */
  fields: IListField[]
  /**
   * Label of the table for assistive technology (the period it lists)
   */
  label: string
  onOpen: (announcement: IAnnouncement) => void
  /**
   * Ref callback for a row, used to scroll a focused announcement into view
   */
  setRowRef?: (id: string) => (element: HTMLElement) => void
}

/**
 * Compact list of announcements: one row each with title, status, level and
 * display period, without the content. A click on a row opens the details.
 */
export const AnnouncementTable: FC<IAnnouncementTableProps> = ({
  announcements,
  fields,
  label,
  onOpen,
  setRowRef
}) => {
  const withTime = (internalName: string) =>
    fields.find((field) => field.internalName === internalName)?.dateOnly === false

  return (
    <div className={styles.table}>
      <Table size='small' aria-label={label} className={styles.grid}>
        <TableHeader>
          <TableRow>
            <TableHeaderCell>{strings.Table.Title}</TableHeaderCell>
            <TableHeaderCell className={styles.narrow}>{strings.Table.Status}</TableHeaderCell>
            <TableHeaderCell className={styles.narrow}>{strings.Table.Level}</TableHeaderCell>
            <TableHeaderCell className={styles.dateColumn}>{strings.Table.From}</TableHeaderCell>
            <TableHeaderCell className={styles.dateColumn}>{strings.Table.To}</TableHeaderCell>
          </TableRow>
        </TableHeader>
        <TableBody>
          {announcements.map((announcement) => (
            <TableRow
              key={announcement.id}
              className={styles.row}
              ref={setRowRef?.(announcement.id)}
              data-announcement-id={announcement.id}
              tabIndex={0}
              aria-haspopup='dialog'
              aria-label={format(strings.Aria.OpenDetails, announcement.title)}
              onClick={() => onOpen(announcement)}
              onKeyDown={(event) => {
                if (event.key === 'Enter' || event.key === ' ') {
                  event.preventDefault()
                  onOpen(announcement)
                }
              }}>
              <TableCell>
                <span className={styles.title}>
                  <IntentIcon announcement={announcement} size={20} />
                  {announcement.pinned && (
                    <Pin16Filled
                      className={styles.pin}
                      title={strings.Aria.PinnedTitle}
                      aria-label={strings.Aria.PinnedTitle}
                    />
                  )}
                  <span>{announcement.title}</span>
                </span>
              </TableCell>
              <TableCell>{statusText(announcement.status)}</TableCell>
              <TableCell>{announcement.values[SEVERITY_FIELD] ?? ''}</TableCell>
              <TableCell className={styles.date}>
                {formatNumericDate(announcement.startDate, withTime('PzlStartDate'))}
              </TableCell>
              <TableCell className={styles.date}>
                {formatNumericDate(announcement.endDate, withTime('PzlEndDate'))}
              </TableCell>
            </TableRow>
          ))}
        </TableBody>
      </Table>
    </div>
  )
}
