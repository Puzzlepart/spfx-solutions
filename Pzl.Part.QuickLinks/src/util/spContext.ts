import { WebPartContext } from '@microsoft/sp-webpart-base'
import { spfi, SPFI, SPFx } from '@pnp/sp'
import { stringIsNullOrEmpty } from '@pnp/common'
import '@pnp/sp/webs'
import '@pnp/sp/lists'
import '@pnp/sp/items'
import '@pnp/sp/batching'
import '@pnp/sp/site-users/web'

let _sp: SPFI = null
let _currentBaseUrl: string = null

export const getSP = (context?: WebPartContext, globalConfigurationUrl?: string): SPFI => {
  if (context) {
    const baseUrl = !stringIsNullOrEmpty(globalConfigurationUrl) ? globalConfigurationUrl : null

    if (_sp === null || _currentBaseUrl !== baseUrl) {
      if (!stringIsNullOrEmpty(globalConfigurationUrl)) {
        _sp = spfi(globalConfigurationUrl).using(SPFx(context))
      } else {
        _sp = spfi().using(SPFx(context))
      }
      _currentBaseUrl = baseUrl
    }
  }
  return _sp
}

export const resetSP = (): void => {
  _sp = null
  _currentBaseUrl = null
}
