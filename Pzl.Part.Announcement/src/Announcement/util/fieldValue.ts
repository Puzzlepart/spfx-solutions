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

const pad = (n: number) => (n < 10 ? `0${n}` : `${n}`)

/**
 * Formats a date for an `<input type="date">` (`yyyy-mm-dd`) or
 * `<input type="datetime-local">` (`yyyy-mm-ddThh:mm`) in local time.
 *
 * @param date Date
 * @param dateOnly Date only
 */
export function toDateInputValue(date: Date, dateOnly: boolean): string {
  const day = `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`
  return dateOnly ? day : `${day}T${pad(date.getHours())}:${pad(date.getMinutes())}`
}

/**
 * Converts a date input value (local time) to an ISO string for the REST API.
 * Returns `null` for empty/invalid input.
 *
 * @param value Input value (`yyyy-mm-dd` or `yyyy-mm-ddThh:mm`)
 */
export function dateInputValueToIso(value: string): string | null {
  if (!value) return null
  const date = new Date(value.length === 10 ? `${value}T00:00` : value)
  return isNaN(date.getTime()) ? null : date.toISOString()
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
