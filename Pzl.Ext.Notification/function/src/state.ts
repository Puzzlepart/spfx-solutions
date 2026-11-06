import { InvocationContext } from '@azure/functions'
import { RestError, TableClient } from '@azure/data-tables'
import { getCredential } from './auth'
import { Settings } from './config'

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

  private async table(): Promise<TableClient> {
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
    if (!this.ready) {
      await this.client.createTable()
      this.ready = true
    }
    return this.client
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
