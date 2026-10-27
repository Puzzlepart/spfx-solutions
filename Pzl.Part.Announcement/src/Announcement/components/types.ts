import { MessageBarIntent } from '@fluentui/react-components'
import { SPUser } from '@microsoft/sp-page-context'
import { WebPartContext } from '@microsoft/sp-webpart-base'
import type { AnnouncementListService } from '../services/AnnouncementListService'

export interface IAnnouncementProps {
  // General
  title: string
  description: string
  targetAudience: boolean

  /**
   * Fields to show in the popover, keyed by internal name (`true` = visible).
   * Fields missing from the map are hidden. Defaults are applied by the
   * web part in `onInit` (see `DEFAULT_POPOVER_FIELDS`).
   */
  popoverFields: Record<string, boolean>
  /**
   * Fields hidden from the create form, keyed by internal name (`false` = hidden).
   * Fields missing from the map are shown if the list schema allows it.
   */
  formFields: Record<string, boolean>

  // Hide/show
  hideHeader: boolean
  /**
   * Hide the "new announcement" button, even for users who may add items to the list
   */
  hideCreateButton: boolean

  // WebPart
  currentUser: SPUser
  context: WebPartContext
  /**
   * Service for the announcement list (schema, permissions, create)
   */
  service: AnnouncementListService
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
  /**
   * Fields of the announcement list (default content type), in content type order
   */
  fields: IListField[]
  /**
   * Whether the current user may add items to the announcement list
   */
  canCreate: boolean
}

export interface IAnnouncement {
  id: string
  title: string
  severity: MessageBarIntent
  content: string
  hasAccess?: boolean
  /**
   * Pinned announcements are sorted first (`PzlPinned`)
   */
  pinned: boolean
  /**
   * Also shown as a banner by the AnnouncementBanner customizer (`PzlShowAsBanner`)
   */
  showAsBanner: boolean
  /**
   * When the Teams notification for this announcement was sent (`PzlNotifiedAt`).
   * Set by the notification service, `null` if not notified.
   */
  notifiedAt?: string
  /**
   * Raw list item values keyed by internal field name. Used to render the
   * fields selected for the popover.
   */
  values: Record<string, any>
}

/**
 * A field of the announcement list, mapped from `SP.Field`
 */
export interface IListField {
  internalName: string
  /**
   * Property name in the REST API (`EntityPropertyName`). Differs from the internal
   * name for fields starting with `_` (`OData__` prefix). Used in `$select`, `$expand`,
   * item values and the payload when creating items.
   */
  entityPropertyName: string
  title: string
  description: string
  /**
   * SharePoint field type (`TypeAsString`), e.g. `Text`, `Note`, `Choice`, `DateTime`, `User`
   */
  type: string
  required: boolean
  readOnly: boolean
  hidden: boolean
  fromBaseType: boolean
  choices?: string[]
  defaultValue?: string
  /**
   * `Note` fields: rich text (HTML)
   */
  richText: boolean
  /**
   * `DateTime` fields: date only (no time part)
   */
  dateOnly: boolean
  /**
   * `User` fields: groups can be selected in addition to people
   */
  allowGroups: boolean
  /**
   * Whether the field is shown in the new form (`ShowInNewForm`)
   */
  showInNewForm: boolean
}
