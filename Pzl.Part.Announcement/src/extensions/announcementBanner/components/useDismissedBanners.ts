import { useState } from 'react'
import { IAnnouncement } from '../../../Announcement/components/types'

/**
 * localStorage key for dismissed banners. Value: `{ [announcement id]: Modified }`,
 * so a dismissed announcement comes back if it is edited.
 */
export const STORAGE_KEY = 'pzl-announcement-banner-dismissed'

/**
 * The value that identifies the version of an announcement a dismissal applies to
 */
function versionOf(announcement: IAnnouncement): string {
  return String(announcement.values?.Modified ?? '')
}

function read(): Record<string, string> {
  try {
    const raw = window.localStorage.getItem(STORAGE_KEY)
    const parsed = raw ? JSON.parse(raw) : {}
    return parsed && typeof parsed === 'object' ? parsed : {}
  } catch {
    return {}
  }
}

function write(dismissed: Record<string, string>): void {
  try {
    window.localStorage.setItem(STORAGE_KEY, JSON.stringify(dismissed))
  } catch {
    // Storage unavailable (private mode, quota): the dismissal lasts for this page only
  }
}

/**
 * Tracks which banners the user has dismissed. Dismissals are kept in
 * localStorage per announcement and version (`Modified`). Entries for
 * announcements that are no longer active are pruned on load.
 *
 * @param announcements Active banner announcements
 */
export function useDismissedBanners(announcements: IAnnouncement[]) {
  const [dismissed, setDismissed] = useState<Record<string, string>>(() => {
    const stored = read()
    const active: Record<string, string> = {}
    announcements.forEach((announcement) => {
      if (stored[announcement.id] === versionOf(announcement)) {
        active[announcement.id] = stored[announcement.id]
      }
    })
    if (Object.keys(active).length !== Object.keys(stored).length) write(active)
    return active
  })

  const isDismissed = (announcement: IAnnouncement) =>
    dismissed[announcement.id] === versionOf(announcement)

  const dismiss = (announcement: IAnnouncement) => {
    const next = { ...dismissed, [announcement.id]: versionOf(announcement) }
    setDismissed(next)
    write(next)
  }

  return { isDismissed, dismiss } as const
}
