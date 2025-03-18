import { MessageBarIntent } from '@fluentui/react-components'
import { SPUser } from '@microsoft/sp-page-context'
import { WebPartContext } from '@microsoft/sp-webpart-base'

export interface IAnnouncementProps {
  // General
  title: string
  description: string
  serverRelativeWebUrl: string
  serviceAnnouncementListUrl: string
  discardForSessionOnly: boolean

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
}

export interface User {
  name: string
  email: string
}
