/**
 * Message level (`PzlSeverity`) and status (`PzlStatus`) of an announcement,
 * mapped from the list's choice values with fixed tables (see CONTEXT.md in
 * Pzl.Part.Announcement). The values are stored in the language the template
 * was applied with, and older lists have the 1.x values. Keep the tables in
 * sync with `Pzl.Part.Announcement/src/Announcement/model.ts`.
 */

export type Severity = 'info' | 'warning' | 'critical'

export type Status = 'planned' | 'ongoing' | 'resolved'

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

const STATUS_BY_VALUE: Record<string, Status> = {
  planlagt: 'planned',
  planned: 'planned',
  pågår: 'ongoing',
  'in progress': 'ongoing',
  ongoing: 'ongoing',
  løst: 'resolved',
  resolved: 'resolved'
}

/** Maps a `PzlSeverity` value. Unknown and empty values count as information. */
export function toSeverity(value: string | null | undefined): Severity {
  return SEVERITY_BY_VALUE[(value ?? '').trim().toLowerCase()] ?? 'info'
}

/**
 * Maps a `PzlSeverity` value from the configuration list («Nivåer som alltid
 * varsles»). Unknown values (free text) are left out instead of counting as
 * information, so a typo never makes every announcement notify.
 */
export function toKnownSeverity(value: string | null | undefined): Severity | null {
  return SEVERITY_BY_VALUE[(value ?? '').trim().toLowerCase()] ?? null
}

/**
 * Maps a `PzlStatus` value. Empty values (lists without the column, items
 * created before it) and unknown values count as ongoing.
 */
export function toStatus(value: string | null | undefined): Status {
  return STATUS_BY_VALUE[(value ?? '').trim().toLowerCase()] ?? 'ongoing'
}
