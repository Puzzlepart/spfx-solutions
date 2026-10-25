/**
 * Settings read from app settings / environment variables.
 * See ../README.md for a description of each setting.
 */
export interface Settings {
  /** Site that hosts the configuration list */
  configSiteUrl: string
  /** Title of the configuration list */
  configListTitle: string
  /** Default title of the announcement list (can be overridden per configuration row) */
  announcementListTitle: string
  /** Severity values (PzlSeverity) that are always notified, regardless of PzlNotifyTeams */
  alwaysNotifySeverities: string[]
  /** How far back to look for new news pages */
  newsLookbackDays: number
  /** Include guest users as recipients */
  includeGuests: boolean
  /**
   * Teams app and entity (static tab) that notifications open. Graph requires `webUrl` to be a
   * Teams deep link, so this is mandatory. Use the notification app itself (its tab opens the
   * intranet) or the intranet's own Teams app.
   */
  teamsDeepLinkAppId: string
  teamsDeepLinkEntityId: string
  /** Storage account used for the state table. Defaults to the function app's own storage account */
  stateStorageAccountName: string
  stateTableName: string
  /** Log what would be sent without sending or writing state */
  dryRun: boolean
}

const env = (name: string, fallback = ''): string => {
  const value = process.env[name]
  return value === undefined || value === null || value === '' ? fallback : value
}

const bool = (value: string): boolean => /^(true|1|yes)$/i.test(value)

export function getSettings(): Settings {
  const configSiteUrl = env('CONFIG_SITE_URL').replace(/\/+$/, '')
  if (!configSiteUrl) throw new Error('App setting CONFIG_SITE_URL is required')

  return {
    configSiteUrl,
    configListTitle: env('CONFIG_LIST_TITLE', 'Varslingskonfigurasjon'),
    announcementListTitle: env('ANNOUNCEMENT_LIST_TITLE', 'Driftsmeldinger'),
    alwaysNotifySeverities: env('ALWAYS_NOTIFY_SEVERITIES', 'Varsel,Alert')
      .split(',')
      .map((s) => s.trim())
      .filter(Boolean),
    newsLookbackDays: Number(env('NEWS_LOOKBACK_DAYS', '7')) || 7,
    includeGuests: bool(env('INCLUDE_GUESTS', 'false')),
    teamsDeepLinkAppId: env('TEAMS_DEEPLINK_APP_ID'),
    teamsDeepLinkEntityId: env('TEAMS_DEEPLINK_ENTITY_ID', 'intranet'),
    stateStorageAccountName: env(
      'STATE_STORAGE_ACCOUNT_NAME',
      env('AzureWebJobsStorage__accountName')
    ),
    stateTableName: env('STATE_TABLE_NAME', 'NotificationState'),
    dryRun: bool(env('DRY_RUN', 'false'))
  }
}
