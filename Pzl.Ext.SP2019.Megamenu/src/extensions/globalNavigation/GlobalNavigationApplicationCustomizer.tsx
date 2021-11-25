import * as React from "react";
import * as ReactDOM from "react-dom";
import { sp } from "@pnp/sp";
import { PnPClientStorage } from "@pnp/common";
import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName,
} from '@microsoft/sp-application-base';
import GlobalNavigation from './GlobalNavigation';
import IGlobalNavigationApplicationCustomizerProperties from "./IGlobalNavigationApplicationCustomizerProperties";
import {
  GlobalNavigationDataFetchBase,
  GlobalNavigationDataFetchJson,
  GlobalNavigationDataFetchSpList
} from "./GlobalNavigationDataFetch";
import {fetchGlobalNavigationSettings} from './SettingsHelper/SettingsHelper';
import * as strings from 'GlobalNavigationApplicationCustomizerStrings';
import ServiceAnnouncement from '../serviceAnnouncement/ServiceAnnouncement';
import "core-js/modules/es6.promise";
import "core-js/modules/es6.array.iterator.js";
import "core-js/modules/es6.array.from.js";
import "whatwg-fetch";
import "es6-map/implement";

const LOG_SOURCE: string = 'GlobalNavigationApplicationCustomizer';

export default class GlobalNavigationApplicationCustomizer extends BaseApplicationCustomizer<IGlobalNavigationApplicationCustomizerProperties> {
  private _topPlaceholder: PlaceholderContent | undefined;
  private child: GlobalNavigation;
  private dataFetch: GlobalNavigationDataFetchBase;

  @override
  public async onInit(): Promise<void> {
    if (this.properties.dataSource) {
      if (this.properties.dataSource.spList) {
        this.dataFetch = new GlobalNavigationDataFetchSpList(this.properties.dataSource.spList);
      } else if (this.properties.dataSource.json) {
        this.dataFetch = new GlobalNavigationDataFetchJson(this.properties.dataSource.json);
      }
    } else {
      Log.info(LOG_SOURCE, 'No data fetch properties specified.');
      return;
    }

    Log.info(LOG_SOURCE, 'onInit');

    let storage = new PnPClientStorage();
    storage.session.deleteExpired();
    storage.local.deleteExpired();
    if (!DEBUG) {
      sp.setup({
        spfxContext: this.context,
        enableCacheExpiration: true,
        defaultCachingStore: "session",
        defaultCachingTimeoutSeconds: 60 * 5,
        globalCacheDisable: false
      });
    } else {
      sp.setup({
        spfxContext: this.context,
        globalCacheDisable: true
      });
    }

    this.context.placeholderProvider.changedEvent.add(this, this.hookMenu);

    return Promise.resolve<void>();
  }

  private async hookMenu() {
    if (!this._topPlaceholder) {
      this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top, { onDispose: this._onDispose });
      if (!this._topPlaceholder) {
        Log.info(LOG_SOURCE, 'The expected placeholder (Top) was not found.');
        return;
      }
      if (this._topPlaceholder.domElement) {
        Log.info(LOG_SOURCE, 'The expected placeholder (Top) was found. Rendering <NavigationContainer />');
        let globalNavigationPlaceholderId = "global-navigation-placeholder";
        let globalNavigationPlaceholder = document.getElementById(globalNavigationPlaceholderId);
        if (globalNavigationPlaceholder == null) {
          globalNavigationPlaceholder = document.createElement("DIV");
          globalNavigationPlaceholder.id = globalNavigationPlaceholderId;
          this._topPlaceholder.domElement.appendChild(globalNavigationPlaceholder);
        }
        const settings = await fetchGlobalNavigationSettings(this.properties.dataSource.spList.serverRelativeWebUrl, this.properties.dataSource.spList.settingsListUrl);
        const globalNavigation = (
          <GlobalNavigation
            ref={instance => { this.child = instance; }}
            dataFetch={this.dataFetch}
            errorText={strings.DefaultLoadErrorText}
            settings={settings}
            currentSiteUrl={this.context.pageContext.site.absoluteUrl}/>
        ); 
        if (!settings.navHideTheMenu || !JSON.parse(settings.navHideTheMenu)) {
          ReactDOM.render(globalNavigation, globalNavigationPlaceholder);
        }
        if (this.properties.serviceAnnouncements) {
          let serviceAnnouncementPlaceholderId = "service-announcement-placeholder";
          let serviceAnnouncementPlaceholder = document.getElementById(serviceAnnouncementPlaceholderId);
          if (serviceAnnouncementPlaceholder == null) {
            serviceAnnouncementPlaceholder = document.createElement("DIV");
            serviceAnnouncementPlaceholder.id = serviceAnnouncementPlaceholderId;
            this._topPlaceholder.domElement.appendChild(serviceAnnouncementPlaceholder);
          }
          // Relying on SP Onlines device detection
          let isMobile = document.getElementsByTagName("body")[0].classList.contains("mobile");
          const serviceAnnouncement = (
            <ServiceAnnouncement
              serverRelativeWebUrl={this.properties.serviceAnnouncements.serverRelativeWebUrl}
              serviceAnnouncementListUrl={this.properties.serviceAnnouncements.listUrl}
              discardForSessionOnly={this.properties.serviceAnnouncements.discardForSessionOnly}
              textAlignment={this.properties.serviceAnnouncements.textAlignment}
              boldText={this.properties.serviceAnnouncements.boldText}
              announcementLevels={settings.announcementLevels}
              isMobile={isMobile} />
          );
          ReactDOM.render(serviceAnnouncement, serviceAnnouncementPlaceholder);
        }
        this.context.application.navigatedEvent.add(this, () => {
          this.child.closeDialog();
        });
      }
    }
  }

  private _onDispose(): void {
    Log.info(LOG_SOURCE, 'Disposed top placeholder.');
  }
}

export { IGlobalNavigationApplicationCustomizerProperties };
