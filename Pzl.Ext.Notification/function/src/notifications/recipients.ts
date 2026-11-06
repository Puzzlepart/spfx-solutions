import { InvocationContext } from '@azure/functions'
import { Settings } from '../config'
import { getGroupMemberIds } from '../graph'
import { SharePointClient } from '../sharepoint'

/**
 * Resolves recipients for an item. Group memberships are cached for the
 * lifetime of the resolver (one function run).
 */
export class RecipientResolver {
  private readonly cache = new Map<string, Promise<string[]>>()

  constructor(
    private readonly settings: Settings,
    private readonly context: InvocationContext
  ) {}

  private members(groupId: string): Promise<string[]> {
    let members = this.cache.get(groupId)
    if (!members) {
      members = getGroupMemberIds(groupId, this.settings.includeGuests).catch((error) => {
        this.context.error(`Could not read members of group ${groupId}: ${error}`)
        return []
      })
      this.cache.set(groupId, members)
    }
    return members
  }

  /**
   * Recipients are the transitive members of the item's audience groups when
   * the item has an audience, otherwise of the default groups for the site.
   *
   * @param sp SharePoint client for the site
   * @param audienceUserIds SharePoint user IDs from `_ModernAudienceTargetUserFieldId`
   * @param defaultGroupIds Entra group IDs from the configuration row
   */
  async resolve(
    sp: SharePointClient,
    audienceUserIds: number[] | undefined,
    defaultGroupIds: string[]
  ): Promise<{ userIds: string[]; groupIds: string[]; source: 'audience' | 'default' }> {
    let groupIds = audienceUserIds?.length ? await sp.resolveGroupIds(audienceUserIds) : []
    const source = groupIds.length ? 'audience' : 'default'
    if (!groupIds.length) groupIds = defaultGroupIds
    const userIds = new Set<string>()
    for (const groupId of groupIds) {
      for (const userId of await this.members(groupId)) userIds.add(userId)
    }
    return { userIds: Array.from(userIds), groupIds, source }
  }
}
