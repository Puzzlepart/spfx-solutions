import { MessageBarIntent } from '@fluentui/react-components'
import { Severity, Status } from '../model'

const INTENT_BY_SEVERITY: Record<Severity, MessageBarIntent> = {
  info: 'info',
  warning: 'warning',
  critical: 'error'
}

/**
 * Message bar intent (colour and icon) for an announcement. A resolved
 * announcement is green whatever its level was, so a red message is never
 * left on the page after the issue is fixed.
 *
 * @param announcement Announcement (message level and status)
 */
export function getIntent(announcement: { severity: Severity; status: Status }): MessageBarIntent {
  if (announcement.status === 'resolved') return 'success'
  return INTENT_BY_SEVERITY[announcement.severity] ?? 'info'
}
