import React, { FC } from 'react'
import {
  FluentProvider,
  IdPrefixProvider,
  Skeleton,
  SkeletonItem,
  Text,
  webLightTheme,
  InfoLabel,
  Button,
  Toaster,
  Toast,
  ToastTitle,
  useToastController
} from '@fluentui/react-components'
import { AddRegular } from '@fluentui/react-icons'
import styles from './Announcement.module.scss'
import { IAnnouncementProps } from './types'
import { AnnouncementContext } from './context'
import { useAnnouncement } from './useAnnouncement'
import { UserMessage } from './UserMessage'
import { AnnouncementDialogs } from './AnnouncementDialogs'
import { getSectionTheme, SectionTheme } from './SectionTheme'
import { CreateAnnouncementDialog } from './CreateAnnouncementDialog'
import { getIntent } from './intent'
import { notifyAnnouncementChanged } from '../util/announcementChanged'
import { format } from '@fluentui/react'
import strings from 'AnnouncementStrings'

export const Announcement: FC<IAnnouncementProps> = (props) => {
  const {
    state,
    setState,
    refresh,
    partId,
    fluentProviderId,
    toasterId,
    openAnnouncement,
    setOpenAnnouncementId,
    setAnnouncementRef,
    popoverFields,
    showCreateButton,
    editor,
    createDialogOpen,
    setCreateDialogOpen
  } = useAnnouncement(props)
  const { dispatchToast } = useToastController(toasterId)

  /**
   * Reads the list again after a change, and tells the banner on the page
   */
  const onChanged = async () => {
    await refresh()
    notifyAnnouncementChanged(partId)
  }

  const onCreated = () => {
    setCreateDialogOpen(false)
    dispatchToast(
      <Toast>
        <ToastTitle>{strings.CreateSuccessText}</ToastTitle>
      </Toast>,
      { intent: 'success' }
    )
    onChanged().catch(() => {
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
              <SectionTheme themeVariant={props.themeVariant} className={styles.header}>
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
                    icon={<AddRegular />}
                    onClick={() => setCreateDialogOpen(true)}>
                    {strings.CreateButtonText}
                  </Button>
                )}
              </SectionTheme>
            )}
            <div
              className={styles.announcements}
              // The focus outline is drawn on the section: its colour follows the section
              style={
                {
                  '--colorStrokeFocus2': getSectionTheme(props.themeVariant).colorStrokeFocus2
                } as React.CSSProperties
              }>
              {state.announcements.length > 0 ? (
                state.announcements.map((announcement) => (
                  <div
                    key={announcement.id}
                    className={styles.message}
                    ref={setAnnouncementRef(announcement.id)}
                    data-announcement-id={announcement.id}
                    role='button'
                    tabIndex={0}
                    aria-haspopup='dialog'
                    aria-label={format(strings.Aria.OpenDetails, announcement.title)}
                    onClick={(event) => {
                      // A link in the content opens the link, not the dialog
                      if ((event.target as HTMLElement).closest?.('a')) return
                      setOpenAnnouncementId(announcement.id)
                    }}
                    onKeyDown={(event) => {
                      if (event.target !== event.currentTarget) return
                      if (event.key === 'Enter' || event.key === ' ') {
                        event.preventDefault()
                        setOpenAnnouncementId(announcement.id)
                      }
                    }}>
                    <UserMessage
                      title={announcement.title}
                      text={announcement.content}
                      intent={getIntent(announcement)}
                      pinned={announcement.pinned}
                      pinnedTitle={strings.Aria.PinnedTitle}
                      status={announcement.status}
                    />
                  </div>
                ))
              ) : (
                <SectionTheme themeVariant={props.themeVariant}>
                  <Text style={{ color: 'var(--colorNeutralForeground4)' }}>
                    {strings.NoAnnouncementsText}
                  </Text>
                </SectionTheme>
              )}
            </div>
          </div>
          <AnnouncementDialogs
            announcement={openAnnouncement}
            fields={popoverFields}
            editor={editor}
            onClose={() => setOpenAnnouncementId(null)}
            onChanged={onChanged}
          />
          {showCreateButton && (
            <CreateAnnouncementDialog
              open={createDialogOpen}
              fields={state.fields}
              hiddenFormFields={props.formFields}
              alwaysNotified={state.alwaysNotified}
              service={props.service}
              onOpenChange={setCreateDialogOpen}
              onSaved={onCreated}
            />
          )}
          <Toaster toasterId={toasterId} position='bottom-end' />
        </FluentProvider>
      </IdPrefixProvider>
    </AnnouncementContext.Provider>
  )
}
