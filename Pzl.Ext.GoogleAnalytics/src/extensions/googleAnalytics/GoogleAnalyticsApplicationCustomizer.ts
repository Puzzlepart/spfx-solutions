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
   *
   *
   * @returns {Promise<void>}
   * @memberof GoogleAnalyticsApplicationCustomizer
   */
  @override
  public onInit(): Promise<void> {
    this.context.application.navigatedEvent.add(this, () => {
      this.track();
    });
    return Promise.resolve();
  }

  /**
   *
   *
   * @private
   * @memberof GoogleAnalyticsApplicationCustomizer
   */
  private track(): void {
    let trackerID: string = this.properties.trackerID;
    if (trackerID) {
      let iframe: any = document.createElement("iframe");
      iframe.id = "GoogleAnalyticsApplicationCustomizerIframe";
      iframe.height = "0";
      iframe.width = "0";
      iframe.style = "display:none;visibility:hidden";
      iframe.src = `https://www.googletagmanager.com/ns.html?id=${trackerID}`;

      let element = document.getElementById("GoogleAnalyticsApplicationCustomizerIframe");
      if (typeof (element) !== 'undefined' && element !== null) {
        document.body.removeChild(element);
      }
      document.body.appendChild(iframe);
      eval(`(function(w,d,s,l,i){w[l]=w[l]||[];w[l].push({'gtm.start':
      new Date().getTime(),event:'gtm.js'});var f=d.getElementsByTagName(s)[0],
      j=d.createElement(s),dl=l!='dataLayer'?'&l='+l:'';j.async=true;j.src=
      'https://www.googletagmanager.com/gtm.js?id='+i+dl;f.parentNode.insertBefore(j,f);
      })(window,document,'script','dataLayer','${trackerID}');
      `);
      Log.info(LOG_SOURCE, `Tracking with ID ${this.properties.trackerID}`);
    } else {
      Log.info(LOG_SOURCE, "Tracking ID not provided");
    }
  }
}
