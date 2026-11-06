/**
 * Message level and status of an announcement (see CONTEXT.md).
 *
 * The list stores choice values as text in the language the template was
 * applied with, and older lists have the 1.x values. The values are mapped
 * with fixed tables, never with the user interface strings, so an announcement
 * looks the same whatever the user's language is. Keep the tables in sync with
 * `Pzl.Ext.Notification/function/src/notifications/levels.ts`.
 */

export type Severity = 'info' | 'warning' | 'critical'

export type Status = 'planned' | 'ongoing' | 'resolved'

/**
 * Choice values of `PzlSeverity` as provisioned by `Templates/Announcement.xml`
 * (Norwegian). Used for the fallback schema and the default level.
 */
export const SEVERITY_VALUES: Record<Severity, string> = {
  info: 'Informasjon',
  warning: 'Advarsel',
  critical: 'Kritisk'
}

/**
 * Choice values of `PzlStatus` as provisioned by `Templates/Announcement.xml`
 * (Norwegian). Used for the fallback schema and the quick actions.
 */
export const STATUS_VALUES: Record<Status, string> = {
  planned: 'Planlagt',
  ongoing: 'Pågår',
  resolved: 'Løst'
}

export const DEFAULT_SEVERITY: Severity = 'warning'

export const DEFAULT_STATUS_VALUE = STATUS_VALUES.planned

/**
 * Known `PzlSeverity` values (lower case) in every language the templates
 * have used, including the 1.x values «Varsel»/«Alert» (now «Kritisk») and
 * «Normal» (now «Informasjon»).
 */
const SEVERITY_BY_VALUE: Record<string, Severity> = {
  informasjon: 'info',
  information: 'info',
  normal: 'info',
  advarsel: 'warning',
  warning: 'warning',
  kritisk: 'critical',
  critical: 'critical',
  varsel: 'critical',
  alert: 'critical'
}

/**
 * Known `PzlStatus` values (lower case) in every language the templates use
 */
const STATUS_BY_VALUE: Record<string, Status> = {
  planlagt: 'planned',
  planned: 'planned',
  pågår: 'ongoing',
  'in progress': 'ongoing',
  ongoing: 'ongoing',
  løst: 'resolved',
  resolved: 'resolved'
}

/**
 * Maps a `PzlSeverity` value to a message level. Unknown and empty values are
 * shown as information.
 */
export function toSeverity(value: string | null | undefined): Severity {
  return SEVERITY_BY_VALUE[(value ?? '').trim().toLowerCase()] ?? 'info'
}

/**
 * Maps a `PzlStatus` value to a status. Empty values (announcements created
 * before the column existed) and unknown values count as ongoing.
 */
export function toStatus(value: string | null | undefined): Status {
  return STATUS_BY_VALUE[(value ?? '').trim().toLowerCase()] ?? 'ongoing'
}

/**
 * Whether a choice value of `PzlSeverity` means the critical level, in any
 * language (used to preselect «Vis som banner» in the create form).
 */
export function isCriticalValue(value: string | null | undefined): boolean {
  return !!value && toSeverity(value) === 'critical'
}

const STATUS_ORDER: Record<Status, number> = { ongoing: 0, planned: 1, resolved: 2 }

const SEVERITY_ORDER: Record<Severity, number> = { critical: 0, warning: 1, info: 2 }

/**
 * Sort order of announcements: pinned first, then status (ongoing, planned,
 * resolved), then message level (critical first), then the newest start date.
 */
export function compareAnnouncements(
  a: { pinned: boolean; status: Status; severity: Severity; startDate?: string },
  b: { pinned: boolean; status: Status; severity: Severity; startDate?: string }
): number {
  return (
    Number(b.pinned) - Number(a.pinned) ||
    STATUS_ORDER[a.status] - STATUS_ORDER[b.status] ||
    SEVERITY_ORDER[a.severity] - SEVERITY_ORDER[b.severity] ||
    (Date.parse(b.startDate ?? '') || 0) - (Date.parse(a.startDate ?? '') || 0)
  )
}

/**
 * The list's choice value for a status: the field's own choice that maps to
 * the status (so a list provisioned in another language gets its own value),
 * or the template's Norwegian value.
 *
 * @param status Status
 * @param choices Choices of the `PzlStatus` field, if known
 */
export function statusChoiceValue(status: Status, choices?: string[]): string {
  return choices?.find((choice) => toStatus(choice) === status) ?? STATUS_VALUES[status]
}
