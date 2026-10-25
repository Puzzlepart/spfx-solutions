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
  Button,
  Toaster,
  Toast,
  ToastTitle,
  useToastController
} from '@fluentui/react-components'
import { Add16Regular } from '@fluentui/react-icons'
import styles from './Announcement.module.scss'
import { IAnnouncementProps } from './types'
import { AnnouncementContext } from './context'
import { useAnnouncement } from './useAnnouncement'
import { UserMessage } from './UserMessage'
import { AnnouncementDetails } from './AnnouncementDetails'
import { CreateAnnouncementDialog } from './CreateAnnouncementDialog'
import { format } from '@fluentui/react'
import strings from 'AnnouncementStrings'

export const Announcement: FC<IAnnouncementProps> = (props) => {
  const {
    state,
    setState,
    refresh,
    fluentProviderId,
    toasterId,
    openAnnouncementId,
    setOpenAnnouncementId,
    setAnnouncementRef,
    toggleAnnouncement,
    popoverFields,
    showCreateButton,
    createDialogOpen,
    setCreateDialogOpen
  } = useAnnouncement(props)
  const { dispatchToast } = useToastController(toasterId)

  const onCreated = () => {
    setCreateDialogOpen(false)
    dispatchToast(
      <Toast>
        <ToastTitle>{strings.CreateSuccessText}</ToastTitle>
      </Toast>,
      { intent: 'success' }
    )
    refresh().catch(() => {
      // Ignore, the list is shown as it was
    })
  }

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
        <FluentProvider theme={webLightTheme} style={{ backgroundColor: 'transparent' }}>
          <div className={styles.announcement}>
            {(!props.hideHeader || showCreateButton) && (
              <div className={styles.header}>
                {!props.hideHeader && props.title && (
                  <Text title={props.title} weight='semibold' size={500} block truncate>
                    {props.title}
                  </Text>
                )}
                {!props.hideHeader && props.description && (
                  <div
                    className={styles.infoLabel}
                    title={format(strings.Aria.HeaderInfoTitle, props.title)}>
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
                {showCreateButton && (
                  <Button
                    className={styles.createButton}
                    appearance='outline'
                    icon={<Add16Regular />}
                    onClick={() => setCreateDialogOpen(true)}>
                    {strings.CreateButtonText}
                  </Button>
                )}
              </div>
            )}
            <div className={styles.announcements}>
              {state.announcements.length > 0 ? (
                state.announcements.map((announcement) => (
                  <Popover
                    key={announcement.id}
                    withArrow
                    closeOnScroll
                    trapFocus
                    positioning='below-start'
                    open={openAnnouncementId === announcement.id}
                    onOpenChange={(_, data) =>
                      setOpenAnnouncementId(data.open ? announcement.id : null)
                    }>
                    <PopoverTrigger>
                      <div
                        className={styles.message}
                        ref={setAnnouncementRef(announcement.id)}
                        data-announcement-id={announcement.id}
                        role='button'
                        tabIndex={0}
                        aria-haspopup='dialog'
                        aria-label={format(strings.Aria.OpenDetails, announcement.title)}
                        onKeyDown={(event) => {
                          if (event.key === 'Enter' || event.key === ' ') {
                            event.preventDefault()
                            toggleAnnouncement(announcement.id)
                          }
                        }}>
                        <UserMessage
                          title={announcement.title}
                          text={announcement.content}
                          intent={announcement.severity}
                          pinned={announcement.pinned}
                          pinnedTitle={strings.Aria.PinnedTitle}
                        />
                      </div>
                    </PopoverTrigger>
                    <PopoverSurface tabIndex={-1}>
                      <AnnouncementDetails announcement={announcement} fields={popoverFields} />
                    </PopoverSurface>
                  </Popover>
                ))
              ) : (
                <Text style={{ color: 'var(--colorNeutralForeground4)' }}>
                  {strings.NoAnnouncementsText}
                </Text>
              )}
            </div>
          </div>
          {showCreateButton && (
            <CreateAnnouncementDialog
              open={createDialogOpen}
              fields={state.fields}
              hiddenFormFields={props.formFields}
              service={props.service}
              onOpenChange={setCreateDialogOpen}
              onCreated={onCreated}
            />
          )}
          <Toaster toasterId={toasterId} position='bottom-end' />
        </FluentProvider>
      </IdPrefixProvider>
    </AnnouncementContext.Provider>
  )
}

Announcement.defaultProps = {
  targetAudience: false,
  hideCreateButton: false
}
