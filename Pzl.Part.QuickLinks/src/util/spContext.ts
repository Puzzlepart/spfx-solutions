import { WebPartContext } from '@microsoft/sp-webpart-base'
import { spfi, SPFI, SPFx } from '@pnp/sp'
import { isNullOrEmpty } from './string'
import '@pnp/sp/webs'
import '@pnp/sp/lists'
import '@pnp/sp/items'
import '@pnp/sp/batching'
import '@pnp/sp/site-users/web'

// One SPFI instance per base URL. A single shared instance breaks when web parts
// with different base URLs render on the same page: each render would recreate the
// instance, and `useEffect([sp])` in the consuming hooks would refetch forever.
const _spInstances = new Map<string, SPFI>()

export const getSP = (context?: WebPartContext, globalConfigurationUrl?: string): SPFI => {
  const baseUrl = !isNullOrEmpty(globalConfigurationUrl) ? globalConfigurationUrl : ''

  if (context && !_spInstances.has(baseUrl)) {
    const sp = baseUrl ? spfi(baseUrl).using(SPFx(context)) : spfi().using(SPFx(context))
    _spInstances.set(baseUrl, sp)
  }
  return _spInstances.get(baseUrl) ?? null
}

export const resetSP = (): void => {
  _spInstances.clear()
}
