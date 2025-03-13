import { SPUser } from '@microsoft/sp-page-context'
import { WebPartContext } from '@microsoft/sp-webpart-base'

export interface IAnnouncementProps {
  // General
  title: string
  description: string

  // WebPart
  hasTeamsContext: boolean
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
  body: string
  created: Date
  modified: Date
}
