import { WebPartContext } from "@microsoft/sp-webpart-base"
import { spfi, SPFI, SPFx } from "@pnp/sp"
import "@pnp/sp/webs"
import "@pnp/sp/lists"
import "@pnp/sp/items"
import "@pnp/sp/batching"
import "@pnp/sp/site-users/web"


let _sp: SPFI = null

export const getSP = (context?: WebPartContext): SPFI => {
  if (_sp === null && context) {
    _sp = spfi().using(SPFx(context))
  }
  return _sp
}
