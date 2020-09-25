import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';

const LOG_SOURCE: string = 'GoogleAnalyticsApplicationCustomizer';
export interface IGoogeAnalyticsApplicationCustomizerProperties {
  trackerID: string;
}

export default class GoogleAnalyticsApplicationCustomizer
  extends BaseApplicationCustomizer<IGoogeAnalyticsApplicationCustomizerProperties> {

  /**
   * @returns {Promise<void>}
   * @memberof GoogleAnalyticsApplicationCustomizer
   */
  @override
  public onInit(): Promise<void> {
    // Is this a virtual pageload?
    if (!(window as any).isNavigatedEventSubscribed) {
      // This is an initial pageload. Attach tracking to navigatedEvent
      this.context.application.navigatedEvent.add(this, this._navigatedEvent);
      (window as any).isNavigatedEventSubscribed = true;
    }
    return Promise.resolve();
  }

  @override
  public onDispose(): Promise<void> {
    this.context.application.navigatedEvent.remove(this, this._navigatedEvent);
    (window as any).isNavigatedEventSubscribed = false;
    (window as any).currentPage = '';
    return Promise.resolve();
  }

  private _navigatedEvent(): void {
    // Timeout for possibility to delay script loading. Avoid if possible.
    setTimeout(() => {
      // Only track if the URL has changed
      if ((window as any).currentPage !== this._getCurrentPage()) {
        this._track();
        (window as any).currentPage = this._getCurrentPage();
      }
    }, 0);
  }

  private _track() {
    const trackerID: string = this.properties.trackerID;
    // Do we have valid-isch tracking code?
    if (trackerID && trackerID.indexOf("GTM-") === 0) {
      // Is GTM instantiated?
      const dataLayerName = 'dataLayer';
      if (typeof((window as any).dataLayer) === "object") {
        // Use the existing GTM dataLayer object to track pageview
        eval(`${dataLayerName}.push({event: 'pageview'});`);
      } else {
        // Instantiate GTM (which will also trigger a pageview)
        eval(`(function(w,d,s,l,i){w[l]=w[l]||[];w[l].push({'gtm.start':
        new Date().getTime(),event:'gtm.js'});var f=d.getElementsByTagName(s)[0],
        j=d.createElement(s),dl=l!='${dataLayerName}'?'&l='+l:'';j.async=true;j.src=
        'https://www.googletagmanager.com/gtm.js?id='+i+dl;f.parentNode.insertBefore(j,f);
        })(window,document,'script','${dataLayerName}','${trackerID}');`);
      }
      Log.info(LOG_SOURCE, `Tracking with ID ${trackerID}`);
    } else if (trackerID.indexOf("UA-") === 0) {
      const gaName = 'ga';
      if (typeof((window as any).ga) === "object") {
        eval(`${gaName}('send', 'pageview');`);
      } else {
        eval(`(function(i,s,o,g,r,a,m){i['GoogleAnalyticsObject']=r;i[r]=i[r]||function(){
          (i[r].q=i[r].q||[]).push(arguments)},i[r].l=1*new Date();a=s.createElement(o),
          m=s.getElementsByTagName(o)[0];a.async=1;a.src=g;m.parentNode.insertBefore(a,m)
          })(window,document,'script','https://www.google-analytics.com/analytics.js','${gaName}');
          ${gaName}('create', '${trackerID}', 'auto');
          ${gaName}('send', 'pageview');`);
      }
      Log.info(LOG_SOURCE, `Tracking with ID ${trackerID}`);
    } else {
      Log.info(LOG_SOURCE, "Tracking ID not provided");
    }
  }

  private _getCurrentPage(): string {
    return window.location.pathname + window.location.search;
  }
}
