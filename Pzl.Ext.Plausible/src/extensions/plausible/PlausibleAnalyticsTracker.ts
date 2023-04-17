import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import Plausible from 'plausible-tracker';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IPlausibleAnalyticsTrackerProperties {
  // This is an example; replace with your own property
  hubSiteId: string;
  apiHost: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class PlausibleAnalyticsTracker extends BaseApplicationCustomizer<IPlausibleAnalyticsTrackerProperties> {

  public onInit(): Promise<void> {
    if (this.properties.hubSiteId && this.properties.hubSiteId !== this.context.pageContext.legacyPageContext.hubSiteId) {
      return Promise.resolve();
    }

    const plausibeHost = this.properties.apiHost ? this.properties.apiHost : 'https://plausible.io';
    const { enableAutoPageviews } = Plausible({ apiHost: plausibeHost });

    enableAutoPageviews();

    return Promise.resolve();
  }
}
