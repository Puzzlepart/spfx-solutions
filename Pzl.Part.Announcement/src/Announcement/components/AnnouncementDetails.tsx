import React, { FC } from 'react'
import { Avatar, Link, Text } from '@fluentui/react-components'
import { Alert16Regular } from '@fluentui/react-icons'
import { format } from '@fluentui/react'
import ReactMarkdown from 'react-markdown'
import rehypeRaw from 'rehype-raw'
import strings from 'AnnouncementStrings'
import styles from './Announcement.module.scss'
import { IAnnouncement, IListField } from './types'
import { CONTENT_FIELD, NOTIFIED_AT_FIELD } from '../constants'
import { formatDate, formatShortDateTime } from '../util/formatDate'
import { isEmptyValue, looksLikeHtml } from '../util/fieldValue'

interface IUserValueProps {
  user: { Title?: string; EMail?: string }
}

const UserValue: FC<IUserValueProps> = ({ user }) => {
  if (!user?.Title) return null
  return (
    <span className={styles.user}>
      <Avatar
        title={user.Title}
        name={user.Title}
        image={{
          src: `/_layouts/15/userphoto.aspx?size=L&accountname=${user.EMail ?? ''}`
        }}
        size={28}
        color='colorful'
        style={{ marginRight: 4 }}
      />
      <span>{user.Title}</span>
    </span>
  )
}

interface IFieldValueProps {
  field: IListField
  value: any
}

/**
 * Renders a list item value based on the field type
 */
const FieldValue: FC<IFieldValueProps> = ({ field, value }) => {
  switch (field.type) {
    case 'Note':
      if (looksLikeHtml(value)) {
        return <div dangerouslySetInnerHTML={{ __html: String(value) }} />
      }
      return field.richText ? (
        <ReactMarkdown rehypePlugins={[rehypeRaw]}>{String(value)}</ReactMarkdown>
      ) : (
        <span style={{ whiteSpace: 'pre-wrap' }}>{String(value)}</span>
      )
    case 'DateTime':
      return <span>{formatDate(String(value), !field.dateOnly)}</span>
    case 'Boolean':
      return <span>{value ? strings.YesText : strings.NoText}</span>
    case 'User':
      return <UserValue user={value} />
    case 'UserMulti':
      return (
        <span className={styles.users}>
          {(value as any[]).map((user, index) => (
            <UserValue key={user.Id ?? index} user={user} />
          ))}
        </span>
      )
    case 'Lookup':
      return <span>{value.Title}</span>
    case 'LookupMulti':
      return <span>{(value as any[]).map((item) => item.Title).join(', ')}</span>
    case 'MultiChoice':
      return <span>{(value as string[]).join(', ')}</span>
    case 'URL':
      return (
        <Link href={value.Url} target='_blank' rel='noreferrer'>
          {value.Description || value.Url}
        </Link>
      )
    case 'TaxonomyFieldType':
      return <span>{value.Label}</span>
    case 'TaxonomyFieldTypeMulti':
      return <span>{(value as any[]).map((term) => term.Label).join(', ')}</span>
    default:
      return <span>{String(value)}</span>
  }
}

export interface IAnnouncementDetailsProps {
  announcement: IAnnouncement
  /**
   * Fields to show, in order
   */
  fields: IListField[]
}

/**
 * Popover content for an announcement: title, the selected fields and the
 * "notified in Teams" footer.
 */
export const AnnouncementDetails: FC<IAnnouncementDetailsProps> = ({ announcement, fields }) => {
  const showNotifiedAt = fields.some((field) => field.internalName === NOTIFIED_AT_FIELD)
  const bodyFields = fields.filter((field) => field.internalName !== NOTIFIED_AT_FIELD)

  return (
    <div className={styles.popover}>
      <Text weight='semibold' size={400} block>
        {announcement.title}
      </Text>
      {bodyFields.map((field) => {
        const value = announcement.values[field.entityPropertyName]
        if (isEmptyValue(value)) return null
        return (
          <div key={field.internalName} className={styles.content}>
            {field.internalName !== CONTENT_FIELD && (
              <Text weight='semibold' size={200} block>
                {field.title}
              </Text>
            )}
            <FieldValue field={field} value={value} />
          </div>
        )
      })}
      {showNotifiedAt && announcement.notifiedAt && (
        <div className={styles.notified}>
          <Alert16Regular />
          <Text size={200}>
            {format(strings.NotifiedAtText, formatShortDateTime(announcement.notifiedAt))}
          </Text>
        </div>
      )}
    </div>
  )
}
