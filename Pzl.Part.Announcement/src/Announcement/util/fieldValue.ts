import { IListField } from '../components/types'

/**
 * Whether a list item value should be treated as empty (not rendered)
 *
 * @param value Field value from the list item
 */
export function isEmptyValue(value: any): boolean {
  if (value === null || value === undefined || value === '') return true
  if (Array.isArray(value)) return value.length === 0
  if (typeof value === 'object') {
    if ('Url' in value) return !value.Url
    if ('Label' in value) return !value.Label
    if ('Title' in value) return !value.Title
  }
  return false
}

/**
 * Whether a string value looks like HTML (rich text stored by SharePoint)
 *
 * @param value Field value
 */
export function looksLikeHtml(value: any): boolean {
  return typeof value === 'string' && /<\/?[a-z][^>]*>/i.test(value)
}

/**
 * Decodes character references (`&#58;`, `&amp;`, `&aring;` …) in a text
 * without tags. SharePoint stores rich text with such references, also for
 * plain characters like the colon.
 *
 * @param text Text without HTML tags
 */
function decodeEntities(text: string): string {
  if (!/&[#\w]+;/.test(text)) return text
  // Parsed as a document and read back as text: nothing is rendered or run
  return new DOMParser().parseFromString(text, 'text/html').documentElement.textContent ?? text
}

/**
 * Text content of an HTML string: tags removed, character references decoded
 * and whitespace collapsed. Used for the banner's preview and to check
 * whether a rich text value is empty.
 *
 * @param html HTML
 */
export function htmlToText(html: string): string {
  return decodeEntities((html ?? '').replace(/<[^>]*>/g, ' '))
    .replace(/[\u200b\ufeff]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
}

const pad = (n: number) => (n < 10 ? `0${n}` : `${n}`)

/**
 * Formats a date as a form value: `yyyy-mm-dd` (date only) or
 * `yyyy-mm-ddThh:mm`, in local time.
 *
 * @param date Date
 * @param dateOnly Date only
 */
export function toDateInputValue(date: Date, dateOnly: boolean): string {
  const day = `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`
  return dateOnly ? day : `${day}T${pad(date.getHours())}:${pad(date.getMinutes())}`
}

/**
 * Parses a form value for a date field (local time). `undefined` for
 * empty/invalid input.
 *
 * @param value Form value (`yyyy-mm-dd` or `yyyy-mm-ddThh:mm`)
 */
export function parseDateInputValue(value: string): Date | undefined {
  if (!value) return undefined
  const date = new Date(value.length === 10 ? `${value}T00:00` : value)
  return isNaN(date.getTime()) ? undefined : date
}

/**
 * Formats a form value for a date field for display: `dd.mm.yyyy`, with
 * `kl. hh:mm` when the value has a time part. Empty for empty/invalid input.
 *
 * @param value Form value (`yyyy-mm-dd` or `yyyy-mm-ddThh:mm`)
 */
export function formatDateInputValue(value: string): string {
  const date = parseDateInputValue(value)
  if (!date) return ''
  const day = `${pad(date.getDate())}.${pad(date.getMonth() + 1)}.${date.getFullYear()}`
  return value.length === 10 ? day : `${day} kl. ${pad(date.getHours())}:${pad(date.getMinutes())}`
}

/**
 * Converts a date input value (local time) to an ISO string for the REST API.
 * Returns `null` for empty/invalid input.
 *
 * @param value Input value (`yyyy-mm-dd` or `yyyy-mm-ddThh:mm`)
 */
export function dateInputValueToIso(value: string): string | null {
  return parseDateInputValue(value)?.toISOString() ?? null
}

/**
 * Initial form value for a field, based on its default value in the list
 *
 * @param field Field
 */
export function getInitialFormValue(field: IListField): any {
  const defaultValue = field.defaultValue
  switch (field.type) {
    case 'Boolean':
      return defaultValue === '1' || /^true$/i.test(defaultValue ?? '')
    case 'DateTime':
      return /^\[today\]$/i.test(defaultValue ?? '')
        ? toDateInputValue(new Date(), field.dateOnly)
        : ''
    case 'MultiChoice':
      return defaultValue ? [defaultValue] : []
    case 'User':
    case 'UserMulti':
      return []
    default:
      return defaultValue ?? ''
  }
}
