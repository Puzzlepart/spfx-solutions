import React, { FC, useState } from 'react'
import {
  Button,
  FluentProvider,
  IdPrefixProvider,
  MessageBar,
  MessageBarActions,
  MessageBarBody,
  MessageBarTitle,
  useId,
  webLightTheme
} from '@fluentui/react-components'
import { CalendarClock20Regular, DismissRegular, Pin16Filled } from '@fluentui/react-icons'
import { format } from '@fluentui/react'
import strings from 'AnnouncementStrings'
import { IAnnouncement, IListField } from '../../../Announcement/components/types'
import { AnnouncementDialog } from '../../../Announcement/components/AnnouncementDialog'
import { htmlToText } from '../../../Announcement/util/fieldValue'
import { getIntent } from '../../../Announcement/components/intent'
import { StatusBadge } from '../../../Announcement/components/StatusBadge'
import { useDismissedBanners } from './useDismissedBanners'
import styles from './AnnouncementBanner.module.scss'

export interface IAnnouncementBannerProps {
  /**
   * Active announcements marked "show as banner", pinned first
   */
  announcements: IAnnouncement[]
  /**
   * Fields shown in the details dialog, in order
   */
  fields: IListField[]
}

/**
 * Maximum length of the content preview shown in the banner
 */
const PREVIEW_LENGTH = 160

function preview(content: string): string {
  const text = htmlToText(content ?? '')
  return text.length > PREVIEW_LENGTH ? `${text.substring(0, PREVIEW_LENGTH - 1)}…` : text
}

/**
 * Message bars at the top of the page, one per announcement. Colour and icon
 * follow the announcement's level and status (see `getIntent`): a resolved
 * announcement is green. Editing an announcement (for example marking it as
 * resolved) changes `Modified`, so a dismissed banner is shown once more. Click opens
 * the same details dialog as the web part (without the editor functions); the
 * close button dismisses the banner for this user.
 */
export const AnnouncementBanner: FC<IAnnouncementBannerProps> = ({ announcements, fields }) => {
  const fluentProviderId = useId('fp-announcement-banner')
  const { isDismissed, dismiss } = useDismissedBanners(announcements)
  const [openId, setOpenId] = useState<string>(null)

  const visible = announcements.filter((announcement) => !isDismissed(announcement))
  if (!visible.length) return null
  const open = visible.find((announcement) => announcement.id === openId)

  return (
    <IdPrefixProvider value={fluentProviderId}>
      <FluentProvider theme={webLightTheme} className={styles.banner}>
        {visible.map((announcement) => (
          <MessageBar
            key={announcement.id}
            intent={getIntent(announcement)}
            icon={announcement.status === 'planned' ? <CalendarClock20Regular /> : undefined}
            layout='multiline'
            className={styles.message}
            role='button'
            tabIndex={0}
            aria-haspopup='dialog'
            aria-label={format(strings.Banner.OpenDetails, announcement.title)}
            onClick={() => setOpenId(announcement.id)}
            onKeyDown={(event) => {
              if (event.key === 'Enter' || event.key === ' ') {
                event.preventDefault()
                setOpenId(announcement.id)
              }
            }}>
            <MessageBarBody>
              <MessageBarTitle>
                {announcement.pinned && (
                  <Pin16Filled
                    className={styles.pin}
                    title={strings.Aria.PinnedTitle}
                    aria-label={strings.Aria.PinnedTitle}
                  />
                )}
                {announcement.title}
                <StatusBadge status={announcement.status} className={styles.status} />
              </MessageBarTitle>{' '}
              {preview(announcement.content)}
            </MessageBarBody>
            <MessageBarActions
              containerAction={
                <Button
                  appearance='transparent'
                  icon={<DismissRegular />}
                  title={strings.Banner.DismissTitle}
                  aria-label={strings.Banner.DismissTitle}
                  onClick={(event) => {
                    event.stopPropagation()
                    dismiss(announcement)
                  }}
                  onKeyDown={(event) => event.stopPropagation()}
                />
              }
            />
          </MessageBar>
        ))}
        <AnnouncementDialog announcement={open} fields={fields} onClose={() => setOpenId(null)} />
      </FluentProvider>
    </IdPrefixProvider>
  )
}
