import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI } from "@pnp/sp";

export interface IMyOwnedSitesProps {
  spfxContext: WebPartContext;
  spClient: SPFI;
  includeSPSites: boolean;
}
