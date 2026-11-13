import { InvocationContext } from '@azure/functions'
import { RestError, TableClient } from '@azure/data-tables'
import { getCredential } from './auth'
import { Settings } from './config'

/** Row that holds when news notifications started for a site (see `newsStart`) */
const NEWS_START_ROW = 'news-start'

/**
 * Sent-state for news pages, stored in Azure Table Storage.
 *
 * Pages are not written to (a field update on a published page creates a new
 * minor version and "unpublished changes"), so the claim lives here instead.
 * Announcements use the `PzlNotifiedAt` column and do not need this table.
 */
export class NotificationState {
  private client: TableClient | undefined
  private ready = false

  constructor(
    private readonly settings: Settings,
    private readonly context: InvocationContext
  ) {}

  /** The table client, without creating the table (for reads, also in a dry run) */
  private tableClient(): TableClient {
    if (!this.client) {
      if (!this.settings.stateStorageAccountName) {
        throw new Error(
          'STATE_STORAGE_ACCOUNT_NAME (or AzureWebJobsStorage__accountName) is required'
        )
      }
      this.client = new TableClient(
        `https://${this.settings.stateStorageAccountName}.table.core.windows.net`,
        this.settings.stateTableName,
        getCredential()
      )
    }
    return this.client
  }

  private async table(): Promise<TableClient> {
    const client = this.tableClient()
    if (!this.ready) {
      await client.createTable()
      this.ready = true
    }
    return client
  }

  /**
   * When news notifications started for a site: the first run that could send for it. News
   * published before that is never notified, so turning the service on, or adding a site, does
   * not notify about the news that is already there.
   *
   * A dry run stores nothing: a site that has not started gets "now" on every run, with
   * `stored: false`.
   */
  async newsStart(siteUrl: string): Promise<{ at: Date; stored: boolean }> {
    const partitionKey = NotificationState.partitionKey(siteUrl)
    const now = new Date()
    if (this.settings.dryRun && !this.settings.stateStorageAccountName) {
      return { at: now, stored: false }
    }
    const stored = await this.readNewsStart(partitionKey)
    if (stored) return { at: stored, stored: true }
    if (this.settings.dryRun) return { at: now, stored: false }
    try {
      await (
        await this.table()
      ).createEntity({ partitionKey, rowKey: NEWS_START_ROW, startedAt: now.toISOString() })
      return { at: now, stored: true }
    } catch (error) {
      // An overlapping run stored it first
      if (error instanceof RestError && error.statusCode === 409) {
        return { at: (await this.readNewsStart(partitionKey)) ?? now, stored: true }
      }
      throw error
    }
  }

  private async readNewsStart(partitionKey: string): Promise<Date | undefined> {
    try {
      const entity = await this.tableClient().getEntity<{ startedAt?: string }>(
        partitionKey,
        NEWS_START_ROW
      )
      const at = entity.startedAt ? new Date(entity.startedAt) : undefined
      return at && !isNaN(at.getTime()) ? at : undefined
    } catch (error) {
      // No row yet, or no table yet
      if (error instanceof RestError && error.statusCode === 404) return undefined
      throw error
    }
  }

  /** Table keys cannot contain / \ # ? */
  static partitionKey(siteUrl: string): string {
    return siteUrl
      .toLowerCase()
      .replace(/^https?:\/\//, '')
      .replace(/[/\#?]/g, '_')
  }

  /**
   * Claims an item for sending. Returns `false` if it was already claimed
   * (sent earlier, or being sent by an overlapping run).
   */
  async claim(
    siteUrl: string,
    rowKey: string,
    properties: Record<string, unknown>
  ): Promise<boolean> {
    if (this.settings.dryRun) return true
    try {
      await (
        await this.table()
      ).createEntity({
        partitionKey: NotificationState.partitionKey(siteUrl),
        rowKey,
        notifiedAt: new Date().toISOString(),
        ...properties
      })
      return true
    } catch (error) {
      if (error instanceof RestError && error.statusCode === 409) return false
      throw error
    }
  }

  /** Marks a claimed item as sent (recipient count for diagnostics) */
  async markSent(siteUrl: string, rowKey: string, recipientCount: number): Promise<void> {
    if (this.settings.dryRun) return
    await (
      await this.table()
    ).updateEntity(
      { partitionKey: NotificationState.partitionKey(siteUrl), rowKey, recipientCount, sent: true },
      'Merge'
    )
  }

  /** Releases a claim after a failed send so the item is retried next run */
  async release(siteUrl: string, rowKey: string): Promise<void> {
    if (this.settings.dryRun) return
    try {
      await (await this.table()).deleteEntity(NotificationState.partitionKey(siteUrl), rowKey)
    } catch (error) {
      this.context.error(`Could not release claim ${rowKey} for ${siteUrl}: ${error}`)
    }
  }
}
