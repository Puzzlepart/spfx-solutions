import { Log } from '@microsoft/sp-core-library';
import { SPComponentLoader } from '@microsoft/sp-loader';

import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import * as strings from 'PzlStylerApplicationCustomizerStrings';

const LOG_SOURCE: string = 'PzlStylerApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IPzlStylerApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class PzlStylerApplicationCustomizer
  extends BaseApplicationCustomizer<IPzlStylerApplicationCustomizerProperties> {

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    SPComponentLoader.loadCss('/sites/CDN/Styling/IntranettStyling.css');

    return Promise.resolve();
  }
}
