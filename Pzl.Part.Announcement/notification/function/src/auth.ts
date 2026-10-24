import { DefaultAzureCredential, TokenCredential } from '@azure/identity'

/**
 * Credential used for Graph, SharePoint and Table Storage.
 *
 * In Azure this resolves to the user-assigned managed identity given by
 * `AZURE_CLIENT_ID`. Locally it falls back to the Azure CLI login or an app
 * registration (`AZURE_TENANT_ID` + `AZURE_CLIENT_ID` + `AZURE_CLIENT_SECRET`).
 */
let credential: TokenCredential | undefined

export function getCredential(): TokenCredential {
  if (!credential) {
    credential = new DefaultAzureCredential({
      managedIdentityClientId: process.env.AZURE_CLIENT_ID
    })
  }
  return credential
}

const tokenCache = new Map<string, { token: string; expiresOn: number }>()

/**
 * Gets an access token for a resource, cached until shortly before it expires.
 *
 * @param scope e.g. `https://graph.microsoft.com/.default` or `https://contoso.sharepoint.com/.default`
 */
export async function getAccessToken(scope: string): Promise<string> {
  const cached = tokenCache.get(scope)
  if (cached && cached.expiresOn - 60_000 > Date.now()) return cached.token
  const result = await getCredential().getToken(scope)
  if (!result) throw new Error(`Could not acquire a token for ${scope}`)
  tokenCache.set(scope, { token: result.token, expiresOn: result.expiresOnTimestamp })
  return result.token
}
