import { MessageBarIntent } from '@fluentui/react-components'
import { Severity } from '../model'

const INTENT_BY_SEVERITY: Record<Severity, MessageBarIntent> = {
  info: 'info',
  warning: 'warning',
  critical: 'error'
}

/**
 * Message bar intent (colour and icon) for an announcement
 *
 * @param announcement Announcement (message level)
 */
export function getIntent(announcement: { severity: Severity }): MessageBarIntent {
  return INTENT_BY_SEVERITY[announcement.severity] ?? 'info'
}
