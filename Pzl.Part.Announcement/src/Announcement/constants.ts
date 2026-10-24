/**
 * Internal name of the content field, rendered as markdown without a label
 */
export const CONTENT_FIELD = 'PzlContent'

/**
 * Internal name of the "notified in Teams" field, rendered as a footer in the popover
 */
export const NOTIFIED_AT_FIELD = 'PzlNotifiedAt'

/**
 * Fields added in v1.1.0 (Teams notification). Lists provisioned with an older
 * template do not have them, so queries are retried without them on failure.
 */
export const NOTIFICATION_FIELDS = ['PzlPinned', 'PzlNotifyTeams', 'PzlNotifiedAt']

/**
 * Default fields shown in the popover (matches the behaviour before the fields
 * became configurable).
 */
export const DEFAULT_POPOVER_FIELDS: Record<string, boolean> = {
  PzlContent: true,
  PzlStartDate: true,
  PzlEndDate: true,
  PzlAffectedSystems: true,
  PzlConsequences: true,
  PzlResponsible: true,
  PzlNotifiedAt: true
}

/**
 * Field types the create form knows how to edit. Other field types are left out
 * of the form (and get their default value from the list).
 */
export const SUPPORTED_FORM_FIELD_TYPES = [
  'Text',
  'Note',
  'Choice',
  'MultiChoice',
  'DateTime',
  'Boolean',
  'User',
  'UserMulti',
  'Number',
  'Currency',
  'URL'
]
