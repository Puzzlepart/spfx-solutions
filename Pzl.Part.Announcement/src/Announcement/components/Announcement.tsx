import React, { FC } from 'react'
import {
  FluentProvider,
  IdPrefixProvider,
  Skeleton,
  SkeletonItem,
  Text,
  webLightTheme,
  InfoLabel,
  Popover,
  PopoverSurface,
  PopoverTrigger,
  Label,
  Avatar
} from '@fluentui/react-components'
import styles from './Announcement.module.scss'
import { IAnnouncementProps } from './types'
import { AnnouncementContext } from './context'
import { useAnnouncement } from './useAnnouncement'
import { UserMessage } from './UserMessage'
import { format } from '@fluentui/react'
import strings from 'AnnouncementStrings'
import { formatDate } from '../util/formatDate'
import ReactMarkdown from 'react-markdown'
import rehypeRaw from 'rehype-raw'

export const Announcement: FC<IAnnouncementProps> = (props) => {
  const { state, setState, fluentProviderId } = useAnnouncement(props)

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
            {!props.hideHeader && (
              <div className={styles.header}>
                {props.description && (
                  <Text title='Driftsmeldinger' weight='semibold' size={500} block truncate>
                    {props.title}
                  </Text>
                )}
                {props.description && (
                  <div
                    className={styles.infoLabel}
                    title={format(strings.Aria.HeaderInfoTitle, props.title)}
                  >
                    <InfoLabel
                      size='medium'
                      info={
                        <div
                          className={styles.infoLabelContent}
                          dangerouslySetInnerHTML={{
                            __html: props.description
                          }}
                        />
                      }
                    />
                  </div>
                )}
              </div>
            )}
            <div className={styles.announcements}>
              {state.announcements.length > 0 ? (
                state.announcements.map((announcement, idx) => (
                  <Popover key={idx} withArrow closeOnScroll positioning='before'>
                    <PopoverTrigger>
                      <div className={styles.message}>
                        <UserMessage
                          key={announcement.id}
                          title={announcement.title}
                          text={announcement.content}
                          intent={announcement.severity}
                        />
                      </div>
                    </PopoverTrigger>
                    <PopoverSurface tabIndex={-1}>
                      <div className={styles.popover}>
                        <Text title='Driftsmeldinger' weight='semibold' size={500} block truncate>
                          {announcement.title}
                        </Text>
                        <div className={styles.content}>
                          <ReactMarkdown rehypePlugins={[rehypeRaw]}>
                            {announcement.content}
                          </ReactMarkdown>
                        </div>
                        <div className={styles.content}>
                          <Label weight='semibold'>Tjenester berørt</Label>
                          <span>{announcement.affectedSystems}</span>
                        </div>
                        <div className={styles.content}>
                          <Label weight='semibold'>Beskrivelse/konsekvens</Label>
                          <span>{announcement.consequence}</span>
                        </div>
                        <div className={styles.content}>
                          <Label weight='semibold'>Ansvarlig</Label>
                          <span>
                            <Avatar
                              title={announcement.responsible.name}
                              name={announcement.responsible.name}
                              image={{
                                src: `/_layouts/15/userphoto.aspx?size=L&accountname=${announcement.responsible.email}`
                              }}
                              size={28}
                              color='colorful'
                              style={{ marginRight: 4 }}
                            />
                            <span>{announcement.responsible.name}</span>
                          </span>
                        </div>
                        <div className={styles.content}>
                          <Label weight='semibold'>Fra dato</Label>
                          <span>{formatDate(announcement.startDate.toString(), true)}</span>
                        </div>
                        <div className={styles.content}>
                          <Label weight='semibold'>Fra dato</Label>
                          <span>{formatDate(announcement.endDate.toString(), true)}</span>
                        </div>
                      </div>
                    </PopoverSurface>
                  </Popover>
                ))
              ) : (
                <Text style={{ color: 'var(--colorNeutralForeground4)' }}>
                  Ingen driftsmeldinger for øyeblikket.
                </Text>
              )}
            </div>
          </div>
        </FluentProvider>
      </IdPrefixProvider>
    </AnnouncementContext.Provider>
  )
}

Announcement.defaultProps = {
  targetAudience: false
}
