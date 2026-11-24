import { getTeamsLoadContentUrl, isTeamsHostOrigin } from './teamsLoad'
import { findAnnouncementIdInUrl, parseAnnouncementId } from './getFocusedAnnouncementId'

const contentUrl =
  'https://contoso.sharepoint.com/_layouts/15/teamslogon.aspx?spfx=true&dest=https%3A%2F%2Fcontoso.sharepoint.com%2F%3Fapp%3Dportals%26driftsmelding%3Ddriftsmelding%3A97'

describe('isTeamsHostOrigin', () => {
  it('accepts the applications that host Teams tabs', () => {
    expect(isTeamsHostOrigin('https://teams.cloud.microsoft')).toBe(true)
    expect(isTeamsHostOrigin('https://teams.microsoft.com')).toBe(true)
    expect(isTeamsHostOrigin('https://outlook.office.com')).toBe(true)
  })

  it('refuses other origins', () => {
    expect(isTeamsHostOrigin('https://contoso.sharepoint.com')).toBe(false)
    expect(isTeamsHostOrigin('http://teams.microsoft.com')).toBe(false)
    expect(isTeamsHostOrigin('https://teams.microsoft.com.example.org')).toBe(false)
    expect(isTeamsHostOrigin('https://notteams.microsoft.com')).toBe(false)
    expect(isTeamsHostOrigin('')).toBe(false)
  })
})

describe('getTeamsLoadContentUrl', () => {
  it('reads contentUrl of a load message', () => {
    const message = { func: 'load', args: [{ entityId: 'tab', contentUrl }], monotonicTimestamp: 1 }
    expect(getTeamsLoadContentUrl(message)).toBe(contentUrl)
  })

  it('ignores other messages', () => {
    expect(getTeamsLoadContentUrl({ func: 'beforeUnload' })).toBeNull()
    expect(getTeamsLoadContentUrl({ id: 56, args: [null] })).toBeNull()
    expect(getTeamsLoadContentUrl({ func: 'load', args: [{}] })).toBeNull()
    expect(getTeamsLoadContentUrl('load')).toBeNull()
    expect(getTeamsLoadContentUrl(null)).toBeNull()
  })
})

describe('parseAnnouncementId', () => {
  it('accepts a bare ID and the deep link prefix', () => {
    expect(parseAnnouncementId('93')).toBe('93')
    expect(parseAnnouncementId('driftsmelding:93')).toBe('93')
  })

  it('ignores empty values and news deep links', () => {
    expect(parseAnnouncementId('')).toBeNull()
    expect(parseAnnouncementId(null)).toBeNull()
    expect(parseAnnouncementId('page:/sites/x/SitePages/a.aspx')).toBeNull()
  })
})

describe('findAnnouncementIdInUrl', () => {
  it('reads the parameter of the address in dest (the Teams tab address)', () => {
    expect(findAnnouncementIdInUrl(contentUrl)).toBe('97')
  })

  it('reads the parameter of the page address', () => {
    expect(findAnnouncementIdInUrl('https://contoso.sharepoint.com/?driftsmelding=12')).toBe('12')
  })

  it('returns null without an announcement ID', () => {
    expect(findAnnouncementIdInUrl('https://contoso.sharepoint.com/?app=portals')).toBeNull()
    expect(findAnnouncementIdInUrl('not a url')).toBeNull()
    expect(
      findAnnouncementIdInUrl(
        'https://contoso.sharepoint.com/_layouts/15/teamslogon.aspx?dest=https%3A%2F%2Fcontoso.sharepoint.com%2F%3Fdriftsmelding%3Dpage%3A%2FSitePages%2Fa.aspx'
      )
    ).toBeNull()
  })
})
