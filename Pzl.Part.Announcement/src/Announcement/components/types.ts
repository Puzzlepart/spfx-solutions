import { MessageBarIntent } from '@fluentui/react-components'
import { SPUser } from '@microsoft/sp-page-context'
import { WebPartContext } from '@microsoft/sp-webpart-base'

export interface IAnnouncementProps {
  // General
  title: string
  description: string
  targetAudience: boolean
  // serverRelativeWebUrl: string
  // serviceAnnouncementListUrl: string
  // discardForSessionOnly: boolean

  // Hide/show
  hideHeader: boolean

  // WebPart
  currentUser: SPUser
  context: WebPartContext
}

export interface IAnnouncementState {
  announcements: IAnnouncement[]
  loading: boolean
  error?: Error
  /**
   * ID of the announcement that should be opened (popover) when the web part loads.
   * Resolved from the Teams deep link (`subPageId`) or the `driftsmelding` query parameter.
   */
  focusedAnnouncementId?: string
}

export interface IAnnouncement {
  id: string
  title: string
  severity: MessageBarIntent
  content: string
  startDate: string
  endDate: string
  affectedSystems: string
  consequence: string
  responsible: User
  hasAccess?: boolean
  /**
   * Pinned announcements are sorted first (`PzlPinned`)
   */
  pinned: boolean
  /**
   * When the Teams notification for this announcement was sent (`PzlNotifiedAt`).
   * Set by the notification service, `null` if not notified.
   */
  notifiedAt?: string
}

export interface User {
  name: string
  email: string
}
