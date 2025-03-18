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
  isMobile: boolean
  textAlignment: Alignment
  boldText: boolean

  // WebPart
  hasTeamsContext: boolean
  currentUser: SPUser
  context: WebPartContext
}

export interface IAnnouncementState {
  announcements: IAnnouncement[]

  modalShouldRender?: boolean
  modalAnnouncement?: IAnnouncement

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
  responsible: {
    name: string
    email: string
  }
}

export enum Alignment {
  Left = 1,
  Center = 2,
  Right = 3
}

export interface User {
  name: string
}
