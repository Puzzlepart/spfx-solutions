import React, { FC, useState } from 'react'
import {
  Button,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  DialogTrigger,
  FluentProvider,
  IdPrefixProvider,
  MessageBar,
  MessageBarActions,
  MessageBarBody,
  MessageBarTitle,
  useId,
  webLightTheme
} from '@fluentui/react-components'
import { Dismiss24Regular, DismissRegular, Pin16Filled } from '@fluentui/react-icons'
import { format } from '@fluentui/react'
import strings from 'AnnouncementStrings'
import { IAnnouncement, IListField } from '../../../Announcement/components/types'
import { AnnouncementDetails } from '../../../Announcement/components/AnnouncementDetails'
import { htmlToText } from '../../../Announcement/util/fieldValue'
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
 * follow the announcement's severity (message bar intent). Click opens a
 * dialog with the same content as the web part's popover; the close button
 * dismisses the banner for this user.
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
            intent={announcement.severity}
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
        <Dialog open={!!open} onOpenChange={(_, data) => !data.open && setOpenId(null)}>
          <DialogSurface>
            {open && (
              <DialogBody>
                <DialogTitle
                  action={
                    <DialogTrigger action='close'>
                      <Button
                        appearance='subtle'
                        icon={<Dismiss24Regular />}
                        aria-label={strings.Aria.CloseDialog}
                      />
                    </DialogTrigger>
                  }>
                  {open.title}
                </DialogTitle>
                <DialogContent className={styles.dialogBody}>
                  <AnnouncementDetails
                    announcement={open}
                    fields={fields}
                    hideTitle
                    style={{ maxWidth: 'none' }}
                  />
                </DialogContent>
                <DialogActions>
                  <DialogTrigger action='close'>
                    <Button appearance='secondary'>{strings.Banner.CloseButtonText}</Button>
                  </DialogTrigger>
                </DialogActions>
              </DialogBody>
            )}
          </DialogSurface>
        </Dialog>
      </FluentProvider>
    </IdPrefixProvider>
  )
}
