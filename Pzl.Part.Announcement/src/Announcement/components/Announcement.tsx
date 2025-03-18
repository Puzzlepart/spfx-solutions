import React, { FC } from 'react'
import {
  FluentProvider,
  IdPrefixProvider,
  Skeleton,
  SkeletonItem,
  Toast,
  ToastBody,
  Toaster,
  ToastTitle,
  Text,
  useToastController,
  webLightTheme
} from '@fluentui/react-components'
import styles from './Announcement.module.scss'
import { IAnnouncementProps } from './types'
import { AnnouncementContext } from './context'
import { useAnnouncement } from './useAnnouncement'
import { UserMessage } from './UserMessage'

export const Announcement: FC<IAnnouncementProps> = (props) => {
  const { state, setState, toasterId, fluentProviderId } = useAnnouncement(props)
  const { dispatchToast } = useToastController(toasterId)

  // dispatchToast(
  //   <Toast appearance='inverted'>
  //     <ToastTitle>Test</ToastTitle>
  //     <ToastBody>Testytest</ToastBody>
  //   </Toast>,
  //   { intent: 'success' }
  // )

  if (state.loading) {
    return (
      <Skeleton>
        <SkeletonItem style={{ width: '192px', height: '40px' }} />
      </Skeleton>
    )
  }

  return (
    <AnnouncementContext.Provider value={{ props, state, setState }}>
      <IdPrefixProvider value={fluentProviderId}>
        <FluentProvider theme={webLightTheme}>
          <div className={styles.announcement}>
            <Text title='Driftsmeldinger' weight='semibold' size={500} block truncate>
              Driftsmeldinger
            </Text>
            <div className={styles.announcements}>
              {state.announcements.length > 0 ? (
                state.announcements.map((announcement, idx) => (
                  <UserMessage
                    key={announcement.id}
                    title={announcement.title}
                    text={announcement.content}
                    intent={announcement.intent}
                  />
                ))
              ) : (
                <Text style={{ color: 'var(--colorNeutralForeground4)' }}>
                  Ingen driftsmeldinger for øyeblikket.
                </Text>
              )}
            </div>
            {/* <Toaster toasterId={toasterId} /> */}
          </div>
        </FluentProvider>
      </IdPrefixProvider>
    </AnnouncementContext.Provider>
  )
}
