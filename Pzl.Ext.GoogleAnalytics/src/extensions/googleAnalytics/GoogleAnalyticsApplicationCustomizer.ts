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
  
  private currentPage = "";
  private isInitialLoad = true;

  private getFreshCurrentPage(): string {
    return window.location.pathname + window.location.search;
  }
  private updateCurrentPage(): void {
    this.currentPage = this.getFreshCurrentPage();
  }

  private navigatedEvent(): void {
    const navigatedPage = this.getFreshCurrentPage();
    if (!this.isInitialLoad && (navigatedPage !== this.currentPage)) {
      this.realNavigatedEvent();
      this.updateCurrentPage();
    }
    this.isInitialLoad = false;
  }

  private realNavigatedEvent(): void {
    this.track();
  }
  /**
   *
   *
   * @returns {Promise<void>}
   * @memberof GoogleAnalyticsApplicationCustomizer
   */
  @override
  public onInit(): Promise<void> {
    this.track();
    this.context.application.navigatedEvent.add(this, () => {
      this.navigatedEvent();
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
      if(trackerID.startsWith("GTM-")) {
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
      } else if (trackerID.startsWith("UA-")) {
         eval(`(function(i,s,o,g,r,a,m){i['GoogleAnalyticsObject']=r;i[r]=i[r]||function(){
          (i[r].q=i[r].q||[]).push(arguments)},i[r].l=1*new Date();a=s.createElement(o),
          m=s.getElementsByTagName(o)[0];a.async=1;a.src=g;m.parentNode.insertBefore(a,m)
          })(window,document,'script','https://www.google-analytics.com/analytics.js','ga');
          ga('create', '${trackerID}', 'auto');
          ga('send', 'pageview');`);
        Log.info(LOG_SOURCE, `Tracking with ID ${this.properties.trackerID}`);
      }
    } else {
      Log.info(LOG_SOURCE, "Tracking ID not provided");
    }
  }
}
