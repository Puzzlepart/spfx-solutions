/**
 * Tells the other parts on the page (the web part and the banner) that an
 * announcement was created or changed, so they read the list again. Both can
 * be on the same page, and each has its own copy of the announcements.
 */
const EVENT_NAME = 'pzl-announcement-changed'

/**
 * Announces a change
 *
 * @param source ID of the part that made the change (it has already refreshed itself)
 */
export function notifyAnnouncementChanged(source: string): void {
  try {
    window.dispatchEvent(new CustomEvent(EVENT_NAME, { detail: { source } }))
  } catch {
    // No CustomEvent: the other parts are updated on the next page load
  }
}

/**
 * Listens for changes made by other parts
 *
 * @param source ID of the listening part; its own changes are ignored
 * @param handler Called when another part has changed an announcement
 * @returns Function that stops listening
 */
export function onAnnouncementChanged(source: string, handler: () => void): () => void {
  const listener = (event: Event) => {
    if ((event as CustomEvent).detail?.source !== source) handler()
  }
  window.addEventListener(EVENT_NAME, listener)
  return () => window.removeEventListener(EVENT_NAME, listener)
}
