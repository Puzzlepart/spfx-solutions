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
