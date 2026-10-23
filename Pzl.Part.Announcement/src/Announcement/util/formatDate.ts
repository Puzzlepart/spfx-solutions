/**
 * Format date
 *
 * @param date Date string or object
 * @param fallback Fallback if date is invalid
 * @param locale Locale
 */
export function formatDate(
  date: string | Date,
  includeTime: boolean = false,
  fallback: string = '',
  locale: string = 'nb-NO'
): string {
  let options: Intl.DateTimeFormatOptions = {
    weekday: 'long',
    year: 'numeric',
    month: 'short',
    day: 'numeric'
  }
  if (includeTime) {
    options = {
      ...options,
      hour: '2-digit',
      minute: '2-digit'
    }
  }
  if (!date) return fallback
  return (typeof date === 'string' ? new Date(date) : date).toLocaleString(locale, options)
}

/**
 * Format date as `dd.mm. kl. hh:mm`, e.g. `23.10. kl. 14:05`.
 * Used for the "notified in Teams" timestamp.
 *
 * @param date Date string or object
 * @param fallback Fallback if date is invalid
 */
export function formatShortDateTime(date: string | Date, fallback: string = ''): string {
  if (!date) return fallback
  const d = typeof date === 'string' ? new Date(date) : date
  if (isNaN(d.getTime())) return fallback
  const pad = (n: number) => (n < 10 ? `0${n}` : `${n}`)
  return `${pad(d.getDate())}.${pad(d.getMonth() + 1)}. kl. ${pad(d.getHours())}:${pad(d.getMinutes())}`
}
