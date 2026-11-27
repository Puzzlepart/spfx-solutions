import React, { FC, Suspense, useEffect, useState } from 'react'
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
import type { IAnnouncementEditor } from '../../../Announcement/components/AnnouncementDialogs'
import {
  notifyAnnouncementChanged,
  onAnnouncementChanged
} from '../../../Announcement/util/announcementChanged'
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
  /**
   * Set for users who may edit the list: the dialog gets «Handlinger og varsling»
   */
  editor?: IAnnouncementEditor
  /**
   * Reads the banner announcements again (after a change)
   */
  reload?: () => Promise<IAnnouncement[]>
}

/**
 * The dialogs with the editor functions (form, date picker, people picker).
 * Fetched only for users who may edit the list, when they open a banner.
 */
const AnnouncementDialogs = React.lazy(() =>
  import(
    /* webpackChunkName: 'announcement-banner-editor' */ '../../../Announcement/components/AnnouncementDialogs'
  ).then((module) => ({ default: module.AnnouncementDialogs }))
)

/**
 * Identifies the banner when it tells the web part on the page about a change
 */
const PART_ID = 'announcement-banner'

/**
 * The most text put in the page for the preview. What is shown is decided by
 * the width: the line is filled and then cut with «…» (see `.body`).
 */
const PREVIEW_LENGTH = 1000

function preview(content: string): string {
  return htmlToText(content ?? '').substring(0, PREVIEW_LENGTH)
}

/**
 * Message bars at the top of the page, one per announcement. Colour and icon
 * follow the announcement's level and status (see `getIntent`): a resolved
 * announcement is green. Editing an announcement (for example marking it as
 * resolved) changes `Modified`, so a dismissed banner is shown once more. Click opens
 * the same details dialog as the web part, with «Handlinger og varsling» for users who
 * may edit the list; the close button dismisses the banner for this user.
 */
export const AnnouncementBanner: FC<IAnnouncementBannerProps> = (props) => {
  const { fields, editor, reload } = props
  const fluentProviderId = useId('fp-announcement-banner')
  const [announcements, setAnnouncements] = useState(props.announcements)
  const { isDismissed, dismiss } = useDismissedBanners(announcements)
  const [openId, setOpenId] = useState<string>(null)

  /**
   * Reads the announcements again; the banners are shown as they were on failure
   */
  const refresh = async () => {
    if (reload) setAnnouncements(await reload())
  }

  // Read the list again when the web part on the page has changed an announcement
  useEffect(
    () =>
      onAnnouncementChanged(PART_ID, () => {
        refresh().catch(() => undefined)
      }),
    []
  )

  const onChanged = async () => {
    await refresh()
    notifyAnnouncementChanged(PART_ID)
  }

  const visible = announcements.filter((announcement) => !isDismissed(announcement))
  const open = visible.find((announcement) => announcement.id === openId)
  const dialog = (
    <AnnouncementDialog announcement={open} fields={fields} onClose={() => setOpenId(null)} />
  )

  return (
    <IdPrefixProvider value={fluentProviderId}>
      <FluentProvider theme={webLightTheme} className={styles.banner}>
        {visible.map((announcement) => (
          <MessageBar
            key={announcement.id}
            intent={getIntent(announcement)}
            icon={announcement.status === 'planned' ? <CalendarClock20Regular /> : undefined}
            layout='multiline'
            className={`${styles.message} ${getIntent(announcement) === 'info' ? styles.info : ''}`}
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
            <MessageBarBody className={styles.body}>
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
        {editor ? (
          // The plain dialog is shown while the editor functions are fetched
          <Suspense fallback={dialog}>
            <AnnouncementDialogs
              announcement={open}
              fields={fields}
              editor={editor}
              onClose={() => setOpenId(null)}
              onChanged={onChanged}
            />
          </Suspense>
        ) : (
          dialog
        )}
      </FluentProvider>
    </IdPrefixProvider>
  )
}
