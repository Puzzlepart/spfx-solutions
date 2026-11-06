/**
 * Internal name of the content field, rendered as markdown without a label
 */
export const CONTENT_FIELD = 'PzlContent'

/**
 * Internal name of the "notified in Teams" field, rendered as a footer in the popover
 */
export const NOTIFIED_AT_FIELD = 'PzlNotifiedAt'

/**
 * Internal name of the "show as banner" field (v1.5.0), read by the
 * AnnouncementBanner application customizer
 */
export const SHOW_AS_BANNER_FIELD = 'PzlShowAsBanner'

/**
 * Internal name of the status field (v2.0.0)
 */
export const STATUS_FIELD = 'PzlStatus'

/**
 * Internal name of the message level field
 */
export const SEVERITY_FIELD = 'PzlSeverity'

/**
 * Fields added after v1.0.0 (pinning in v1.1.0, banner in v1.5.0, status in
 * v2.0.0) and the fields of the notification add-on. Lists provisioned with an
 * older template do not have them, so queries are retried without them on
 * failure.
 */
export const OPTIONAL_FIELDS = [
  'PzlPinned',
  SHOW_AS_BANNER_FIELD,
  STATUS_FIELD,
  'PzlNotifyTeams',
  'PzlNotifyAgain',
  NOTIFIED_AT_FIELD
]

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
 * Fields written by the notification service, or only meaningful on existing
 * items («Varsle på nytt»). Never shown in the create form, even if the list
 * schema does not say `ShowInNewForm="FALSE"`.
 */
export const SERVICE_MANAGED_FIELDS = [NOTIFIED_AT_FIELD, 'PzlNotifyAgain']

/**
 * Fields hidden in the create form unless turned on in the web part
 * properties. «Varsle i Teams» belongs to the notification add-on and is left
 * on lists that once had it, so it is opt-in.
 */
export const DEFAULT_HIDDEN_FORM_FIELDS = ['PzlNotifyTeams']

/**
 * Whether a field is shown in the create form, given the web part's
 * `formFields` setting (`true` = shown, `false` = hidden, missing = default)
 *
 * @param internalName Internal name of the field
 * @param formFields The web part's `formFields` property
 */
export function isFormFieldEnabled(
  internalName: string,
  formFields: Record<string, boolean> | undefined
): boolean {
  const setting = formFields?.[internalName]
  if (typeof setting === 'boolean') return setting
  return !DEFAULT_HIDDEN_FORM_FIELDS.includes(internalName)
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
