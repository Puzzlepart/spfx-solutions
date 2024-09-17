import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import { SPComponentLoader } from '@microsoft/sp-loader';

import * as strings from 'PzlStylerApplicationCustomizerStrings';

const LOG_SOURCE: string = 'PzlStylerApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IPzlStylerApplicationCustomizerProperties {  
  cssFilePath: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class PzlStylerApplicationCustomizer
  extends BaseApplicationCustomizer<IPzlStylerApplicationCustomizerProperties> {

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);
    if (this.properties.cssFilePath) {
      let cssFileToLoad = this.properties.cssFilePath.trim();
      if (this.properties.cssFilePath.indexOf('https://') === -1) {
        if (this.properties.cssFilePath.indexOf('/') === 0) {
          cssFileToLoad = `${document.location.protocol}//${document.location.hostname}${this.properties.cssFilePath.trim()}`;          
        } else {
          cssFileToLoad = `${document.location.protocol}//${document.location.hostname}/${this.properties.cssFilePath.trim()}`;
        }
      }
      SPComponentLoader.loadCss(cssFileToLoad);
    }

    return Promise.resolve();
  }
}
