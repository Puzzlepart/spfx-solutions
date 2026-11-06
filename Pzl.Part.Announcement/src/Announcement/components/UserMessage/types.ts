import { MessageBarProps } from '@fluentui/react-components'
import { Status } from '../../model'

/**
 * @category UserMessage
 */
export interface IUserMessageProps extends MessageBarProps {
  /**
   * Title to show in the component, should be a descriptive title of the message
   */
  title?: string

  /**
   * Text to show in the component
   *
   * NOTE: Supports markdown and HTML
   */
  text?: string

  /**
   * Show a pin icon in front of the title
   */
  pinned?: boolean

  /**
   * Title (tooltip) for the pin icon
   */
  pinnedTitle?: string

  /**
   * Status of the announcement. Planned and resolved announcements get a
   * label after the title, planned ones a calendar icon.
   */
  status?: Status

  /**
   * On click handler for the message
   */
  onClick?: (event: React.MouseEvent<any>) => void

  /**
   * Container style
   */
  containerStyle?: React.CSSProperties

  /**
   * To flex the message center, speficy a min height
   */
  fixedCenter?: number

  /**
   * To reduce size of the message bar and make it compact
   */
  isCompact?: boolean

  /**
   * Link target for markdown links
   */
  linkTarget?: string
}
