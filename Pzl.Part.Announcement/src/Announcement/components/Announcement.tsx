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
import { AnnouncementDialog } from './AnnouncementDialog'
import { NotificationSettings } from './NotificationSettings'
import { CreateAnnouncementDialog } from './CreateAnnouncementDialog'
import { getIntent } from './intent'
import { EditAction, StatusActions } from './QuickActions'
import { statusText } from './StatusBadge'
import { IAnnouncement } from './types'
import { Status } from '../model'
import { format } from '@fluentui/react'
import strings from 'AnnouncementStrings'

export const Announcement: FC<IAnnouncementProps> = (props) => {
  const {
    state,
    setState,
    refresh,
    fluentProviderId,
    toasterId,
    openAnnouncement,
    setOpenAnnouncementId,
    dialogError,
    setAnnouncementRef,
    popoverFields,
    showCreateButton,
    showQuickActions,
    setStatus,
    editAnnouncement,
    openCreateDialog,
    openEditDialog,
    showNotificationSettings,
    notifyTeamsField,
    notifyAgainField,
    openAlwaysNotified,
    setNotification,
    createDialogOpen,
    setCreateDialogOpen
  } = useAnnouncement(props)
  const { dispatchToast } = useToastController(toasterId)

  const onSaved = () => {
    setCreateDialogOpen(false)
    dispatchToast(
      <Toast>
        <ToastTitle>
          {editAnnouncement ? strings.EditSuccessText : strings.CreateSuccessText}
        </ToastTitle>
      </Toast>,
      { intent: 'success' }
    )
    refresh().catch(() => {
      // Ignore, the list is shown as it was
    })
  }

  /**
   * Changes the status. The dialog closes on success; a failure is shown in
   * the dialog.
   */
  const onSetStatus = async (announcement: IAnnouncement, status: Status) => {
    await setStatus(announcement, status)
    dispatchToast(
      <Toast>
        <ToastTitle>
          {format(strings.QuickActions.StatusChangedText, statusText(status).toLowerCase())}
        </ToastTitle>
      </Toast>,
      { intent: 'success' }
    )
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
                    icon={<AddRegular />}
                    onClick={openCreateDialog}>
                    {strings.CreateButtonText}
                  </Button>
                )}
              </div>
            )}
            <div className={styles.announcements}>
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
                <Text style={{ color: 'var(--colorNeutralForeground4)' }}>
                  {strings.NoAnnouncementsText}
                </Text>
              )}
            </div>
          </div>
          <AnnouncementDialog
            announcement={openAnnouncement}
            fields={popoverFields}
            error={dialogError}
            onClose={() => setOpenAnnouncementId(null)}
            startActions={
              state.canEdit &&
              openAnnouncement && <EditAction onEdit={() => openEditDialog(openAnnouncement)} />
            }
            endActions={
              showQuickActions &&
              openAnnouncement && (
                <StatusActions announcement={openAnnouncement} onSetStatus={onSetStatus} />
              )
            }>
            {showNotificationSettings && openAnnouncement && (
              <NotificationSettings
                announcement={openAnnouncement}
                notifyTeamsField={notifyTeamsField}
                notifyAgainField={notifyAgainField}
                alwaysNotified={openAlwaysNotified}
                onChange={setNotification}
              />
            )}
          </AnnouncementDialog>
          {(showCreateButton || state.canEdit) && (
            <CreateAnnouncementDialog
              open={createDialogOpen}
              fields={state.fields}
              hiddenFormFields={props.formFields}
              alwaysNotified={state.alwaysNotified}
              service={props.service}
              onOpenChange={setCreateDialogOpen}
              announcement={editAnnouncement ?? undefined}
              onSaved={onSaved}
            />
          )}
          <Toaster toasterId={toasterId} position='bottom-end' />
        </FluentProvider>
      </IdPrefixProvider>
    </AnnouncementContext.Provider>
  )
}
