import { Log } from '@microsoft/sp-core-library';
import { SPComponentLoader } from '@microsoft/sp-loader';

import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import * as strings from 'PolcssApplicationCustomizerStrings';

const LOG_SOURCE: string = 'PolcssApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IPolcssApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class PolcssApplicationCustomizer
  extends BaseApplicationCustomizer<IPolcssApplicationCustomizerProperties> {

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    SPComponentLoader.loadCss('https://folkehelse.sharepoint.com/sites/CDN/Styling/IntranettStyling.css');

    let message: string = this.properties.testMessage;
    if (!message) {
      message = '(No properties were provided.)';
    }

    return Promise.resolve();
  }
}
