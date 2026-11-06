import { getAccessToken } from './auth'
import { HttpError, request } from './http'

/**
 * SharePoint REST client (app-only, `Sites.Selected`).
 *
 * Tokens are acquired for the tenant of the site, e.g. `https://contoso.sharepoint.com/.default`.
 */
export class SharePointClient {
  constructor(public readonly siteUrl: string) {}

  private get resource(): string {
    const { origin } = new URL(this.siteUrl)
    return `${origin}/.default`
  }

  private async headers(extra: Record<string, string> = {}): Promise<Record<string, string>> {
    return {
      Authorization: `Bearer ${await getAccessToken(this.resource)}`,
      Accept: 'application/json;odata=nometadata',
      ...extra
    }
  }

  /** Server-relative URL of the web, e.g. `/sites/intranet` or `/` */
  get serverRelativeUrl(): string {
    const { pathname } = new URL(this.siteUrl)
    return pathname.replace(/\/+$/, '') || '/'
  }

  async get<T>(path: string): Promise<T> {
    return request<T>(`${this.siteUrl}/_api/${path}`, { headers: await this.headers() })
  }

  /**
   * Queries list items. Fields in `optionalSelect` are dropped and the query
   * retried if SharePoint rejects them (e.g. audience targeting not enabled).
   */
  async getItems<T>(
    listPath: string,
    query: {
      select: string[]
      optionalSelect?: string[]
      filter?: string
      top?: number
      orderBy?: string
    }
  ): Promise<T[]> {
    const build = (select: string[]) => {
      const params = [`$select=${select.join(',')}`, `$top=${query.top ?? 500}`]
      if (query.filter) params.push(`$filter=${encodeURIComponent(query.filter)}`)
      if (query.orderBy) params.push(`$orderby=${encodeURIComponent(query.orderBy)}`)
      return `${listPath}/items?${params.join('&')}`
    }
    try {
      const result = await this.get<{ value: T[] }>(
        build([...query.select, ...(query.optionalSelect ?? [])])
      )
      return result.value
    } catch (error) {
      if (query.optionalSelect?.length && error instanceof HttpError && error.status === 400) {
        const result = await this.get<{ value: T[] }>(build(query.select))
        return result.value
      }
      throw error
    }
  }

  /** Updates fields on a list item (MERGE) */
  async updateItem(listPath: string, id: number, fields: Record<string, unknown>): Promise<void> {
    await request(`${this.siteUrl}/_api/${listPath}/items(${id})`, {
      method: 'POST',
      headers: await this.headers({
        'X-HTTP-Method': 'MERGE',
        'IF-MATCH': '*',
        'Content-Type': 'application/json;odata=nometadata'
      }),
      body: fields
    })
  }

  /** `web/lists/getbytitle('...')` path for a list title */
  static listByTitle(title: string): string {
    return `web/lists/getbytitle('${title.replace(/'/g, "''")}')`
  }

  /** `web/GetList('...')` path for a server-relative list URL */
  static listByUrl(serverRelativeUrl: string): string {
    return `web/GetList('${serverRelativeUrl.replace(/'/g, "''")}')`
  }

  /**
   * Resolves SharePoint user IDs (e.g. from `_ModernAudienceTargetUserFieldId`)
   * to Entra group object IDs. Entries that are not groups are ignored.
   */
  async resolveGroupIds(userIds: number[]): Promise<string[]> {
    const groupIds: string[] = []
    for (const id of userIds) {
      const user = await this.get<{ LoginName: string }>(`web/getuserbyid(${id})?$select=LoginName`)
      const groupId = parseGroupId(user.LoginName)
      if (groupId) groupIds.push(groupId)
    }
    return groupIds
  }
}

/**
 * Extracts the Entra group ID from a SharePoint claim login name:
 * `c:0t.c|tenant|<guid>` (security group), `c:0o.c|federateddirectoryclaimprovider|<guid>`
 * (Microsoft 365 group) or `..|<guid>_o` (group owners).
 */
export function parseGroupId(loginName: string): string | null {
  const match = /\|([0-9a-f-]{36})(_o)?$/i.exec(loginName ?? '')
  return match ? match[1].toLowerCase() : null
}

/** OData datetime literal for `$filter` */
export function odataDate(date: Date): string {
  return `datetime'${date.toISOString()}'`
}
